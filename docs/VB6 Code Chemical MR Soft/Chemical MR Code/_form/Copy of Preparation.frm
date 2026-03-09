VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form frmPreparation 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Chemical MR"
   ClientHeight    =   11880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
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
   ScaleHeight     =   11880
   ScaleWidth      =   19080
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
         Left            =   120
         ScaleHeight     =   50002.14
         ScaleMode       =   0  'User
         ScaleWidth      =   19155
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   19155
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
            Height          =   7560
            Index           =   5
            Left            =   960
            TabIndex        =   199
            Top             =   31320
            Visible         =   0   'False
            Width           =   17175
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   26
               Left            =   11040
               TabIndex        =   211
               Top             =   6000
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Select Bottle"
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
                  TabIndex        =   212
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
               Index           =   25
               Left            =   14160
               TabIndex        =   209
               Top             =   6000
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Goto Preparation"
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
                  Index           =   25
                  Left            =   0
                  TabIndex        =   210
                  Top             =   120
                  Width           =   3015
               End
            End
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
               Index           =   3
               Left            =   0
               TabIndex        =   206
               Top             =   0
               Width           =   17175
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mother Solution"
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
                  Index           =   4
                  Left            =   7695
                  TabIndex        =   208
                  Top             =   105
                  Width           =   1875
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Select mother solution from table"
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
                  Left            =   14100
                  TabIndex        =   207
                  Top             =   120
                  Width           =   2985
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   24
               Left            =   0
               TabIndex        =   204
               Top             =   6000
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "All MR Bottle"
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
                  TabIndex        =   205
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frMotherTable 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   1335
               Left            =   6120
               TabIndex        =   202
               Top             =   2040
               Visible         =   0   'False
               Width           =   5055
               Begin VB.Label lbMotherTable 
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
                  Height          =   495
                  Left            =   0
                  TabIndex        =   203
                  Top             =   525
                  Width           =   4995
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   23
               Left            =   7920
               TabIndex        =   200
               Top             =   6000
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Close Bottle"
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
                  TabIndex        =   201
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin FlexCell.Grid Grid5 
               Height          =   4935
               Left            =   0
               TabIndex        =   213
               TabStop         =   0   'False
               Top             =   720
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
            Begin VB.Line Line2 
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   17280
               Y1              =   5760
               Y2              =   5760
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
            Height          =   7560
            Index           =   6
            Left            =   3360
            TabIndex        =   220
            Top             =   30240
            Visible         =   0   'False
            Width           =   17175
            Begin VB.Frame Frame1 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   1335
               Left            =   6120
               TabIndex        =   226
               Top             =   2040
               Visible         =   0   'False
               Width           =   5055
               Begin VB.Label Label12 
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
                  Height          =   495
                  Left            =   0
                  TabIndex        =   227
                  Top             =   525
                  Width           =   4995
               End
            End
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
               Index           =   6
               Left            =   0
               TabIndex        =   223
               Top             =   0
               Width           =   17175
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Select Pipette from table"
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
                  Left            =   14820
                  TabIndex        =   225
                  Top             =   120
                  Width           =   2265
               End
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pipette"
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
                  Index           =   6
                  Left            =   8205
                  TabIndex        =   224
                  Top             =   105
                  Width           =   855
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   27
               Left            =   14160
               TabIndex        =   221
               Top             =   6000
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Exit"
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
                  Index           =   27
                  Left            =   0
                  TabIndex        =   222
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin FlexCell.Grid Grid6 
               Height          =   4935
               Left            =   0
               TabIndex        =   228
               TabStop         =   0   'False
               Top             =   720
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
            Begin VB.Line Line7 
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   17280
               Y1              =   5760
               Y2              =   5760
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Caption         =   "&H00E0E0E0&"
            Height          =   9015
            Index           =   4
            Left            =   960
            TabIndex        =   136
            Top             =   40798
            Visible         =   0   'False
            Width           =   17055
            Begin VB.TextBox txRevision 
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
               TabIndex        =   162
               Top             =   5760
               Width           =   2415
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
               Left            =   6000
               Style           =   2  'Dropdown List
               TabIndex        =   161
               Top             =   5760
               Visible         =   0   'False
               Width           =   2415
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00886010&
               BorderStyle     =   0  'None
               Height          =   1335
               Left            =   5880
               TabIndex        =   147
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
                  TabIndex        =   149
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
                  TabIndex        =   148
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
               TabIndex        =   145
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
                  TabIndex        =   146
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
               TabIndex        =   142
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
                  TabIndex        =   144
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
                  TabIndex        =   143
                  Top             =   75
                  Width           =   2055
               End
            End
            Begin VB.TextBox txRevision 
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
               TabIndex        =   141
               Top             =   5760
               Width           =   2415
            End
            Begin VB.TextBox txRevision 
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
               TabIndex        =   140
               Top             =   5760
               Width           =   5655
            End
            Begin VB.TextBox txRevision 
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
               TabIndex        =   139
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
               TabIndex        =   137
               Top             =   6960
               Width           =   3015
               Begin VB.Image Image1 
                  Height          =   480
                  Left            =   120
                  MousePointer    =   99  'Custom
                  OLEDropMode     =   1  'Manual
                  Picture         =   "Preparation.frx":29F2
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
                  TabIndex        =   138
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin FlexCell.Grid Grid4 
               Height          =   4695
               Left            =   0
               TabIndex        =   150
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
               Index           =   4
               Left            =   8400
               TabIndex        =   151
               Top             =   7440
               Width           =   1815
            End
            Begin VB.Label lbFunction 
               BackStyle       =   0  'Transparent
               Height          =   975
               Index           =   5
               Left            =   6720
               TabIndex        =   152
               Top             =   7320
               Width           =   1695
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
               TabIndex        =   163
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
               TabIndex        =   158
               Top             =   8640
               Width           =   6435
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   4
               Left            =   9120
               Picture         =   "Preparation.frx":5DD4
               ToolTipText     =   "4000"
               Top             =   7440
               Width           =   240
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   5
               Left            =   7320
               Picture         =   "Preparation.frx":67D6
               ToolTipText     =   "4000"
               Top             =   7440
               Width           =   240
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Save Note"
               ForeColor       =   &H00808080&
               Height          =   255
               Left            =   7080
               TabIndex        =   157
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
               TabIndex        =   156
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
               TabIndex        =   155
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
               TabIndex        =   154
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
               TabIndex        =   153
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8655
            Index           =   2
            Left            =   1080
            TabIndex        =   26
            Top             =   20400
            Visible         =   0   'False
            Width           =   17295
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   19
               Left            =   120
               TabIndex        =   181
               Top             =   5880
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Select Bottle"
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
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   24
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00886010&
               Height          =   630
               Index           =   0
               Left            =   5880
               Locked          =   -1  'True
               TabIndex        =   167
               Text            =   "245,456"
               Top             =   2340
               Width           =   2175
            End
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   24
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00886010&
               Height          =   585
               Index           =   1
               Left            =   10680
               Locked          =   -1  'True
               TabIndex        =   166
               Text            =   "34F"
               Top             =   2340
               Width           =   1455
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
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
               ForeColor       =   &H00000000&
               Height          =   400
               Index           =   19
               Left            =   12960
               Locked          =   -1  'True
               TabIndex        =   164
               Text            =   "mg/L"
               Top             =   4320
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
               Index           =   18
               Left            =   6120
               Locked          =   -1  'True
               TabIndex        =   134
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
               TabIndex        =   115
               Top             =   6480
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Scan QRCode Bottle"
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
                  TabIndex        =   116
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
               Left            =   3360
               TabIndex        =   107
               Top             =   5040
               Width           =   10815
            End
            Begin VB.PictureBox PicTolerance 
               BorderStyle     =   0  'None
               Height          =   135
               Left            =   6720
               ScaleHeight     =   135
               ScaleWidth      =   4095
               TabIndex        =   106
               Top             =   4560
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
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
               ForeColor       =   &H00000000&
               Height          =   400
               Index           =   16
               Left            =   12960
               Locked          =   -1  'True
               TabIndex        =   104
               Text            =   "0,40"
               Top             =   3840
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
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
               ForeColor       =   &H00000000&
               Height          =   400
               Index           =   15
               Left            =   12960
               Locked          =   -1  'True
               TabIndex        =   102
               Text            =   "1"
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00886010&
               Height          =   400
               Index           =   14
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   100
               Text            =   "-200,221"
               Top             =   3000
               Visible         =   0   'False
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
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   98
               Text            =   "-21 %"
               Top             =   4080
               Visible         =   0   'False
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
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   96
               Text            =   "-200,221"
               Top             =   3600
               Visible         =   0   'False
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
               TabIndex        =   94
               Text            =   "1229,998"
               Top             =   3960
               Width           =   4095
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
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
               ForeColor       =   &H00000000&
               Height          =   400
               Index           =   10
               Left            =   6720
               Locked          =   -1  'True
               TabIndex        =   92
               Text            =   "1300,400"
               Top             =   3360
               Width           =   4095
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
               Left            =   14040
               Locked          =   -1  'True
               TabIndex        =   90
               Top             =   1800
               Width           =   2415
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   9
               Left            =   14280
               TabIndex        =   79
               Top             =   6480
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pipette Selection"
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
                  TabIndex        =   80
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00307030&
               BorderStyle     =   0  'None
               Height          =   495
               Index           =   8
               Left            =   3240
               TabIndex        =   77
               Top             =   5880
               Visible         =   0   'False
               Width           =   10935
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "ACQUIRE"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   12
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
                  TabIndex        =   78
                  Top             =   120
                  Width           =   10935
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   7
               Left            =   14280
               TabIndex        =   75
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
                  TabIndex        =   76
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
               Left            =   1920
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
               Index           =   1
               Left            =   6120
               Locked          =   -1  'True
               TabIndex        =   73
               Top             =   1080
               Width           =   2655
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
               Left            =   14640
               Locked          =   -1  'True
               TabIndex        =   72
               Top             =   1080
               Width           =   1815
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
               Left            =   10560
               Locked          =   -1  'True
               TabIndex        =   71
               Top             =   1080
               Width           =   1815
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
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   70
               Top             =   1440
               Width           =   855
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
               Left            =   10560
               Locked          =   -1  'True
               TabIndex        =   69
               Top             =   1440
               Width           =   1815
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
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   68
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
               Left            =   6120
               Locked          =   -1  'True
               TabIndex        =   67
               Top             =   1800
               Width           =   2655
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
               Left            =   10560
               Locked          =   -1  'True
               TabIndex        =   66
               Top             =   1800
               Width           =   1800
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
               Height          =   780
               Index           =   2
               Left            =   120
               TabIndex        =   27
               Top             =   0
               Width           =   17295
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "STD : Acquisition Details"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   24
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   585
                  Index           =   2
                  Left            =   5910
                  TabIndex        =   29
                  Top             =   40
                  Width           =   5415
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
                  Top             =   240
                  Width           =   1050
               End
            End
            Begin VB.Label lbSTDNote_AQ 
               Alignment       =   2  'Center
               BackColor       =   &H00644603&
               Caption         =   "Label13"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   120
               TabIndex        =   230
               Top             =   5520
               Visible         =   0   'False
               Width           =   17175
            End
            Begin VB.Shape Shape3 
               BorderColor     =   &H00C0C0C0&
               Height          =   855
               Left            =   1920
               Top             =   2280
               Width           =   14535
            End
            Begin VB.Label lbStock 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bottle Qty LEFT"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   4080
               TabIndex        =   170
               Top             =   2520
               Width           =   1380
            End
            Begin VB.Label lbStock 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bottle #"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   9840
               TabIndex        =   169
               Top             =   2520
               Width           =   720
            End
            Begin VB.Label lbUM 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "mg/L"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   8280
               TabIndex        =   168
               Top             =   2520
               Width           =   495
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "STD Unit"
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
               Index           =   19
               Left            =   12060
               TabIndex        =   165
               Top             =   4320
               Width           =   705
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "MR value ( concentration )"
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
               Left            =   3600
               TabIndex        =   135
               Top             =   1440
               Width           =   2385
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
               TabIndex        =   129
               Top             =   7200
               Width           =   17235
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
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
               Left            =   2160
               TabIndex        =   108
               Top             =   5040
               Width           =   1095
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "STD Value"
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
               Left            =   11880
               TabIndex        =   105
               Top             =   3840
               Width           =   885
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "STD Number"
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
               Left            =   11700
               TabIndex        =   103
               Top             =   3360
               Width           =   1065
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qty Left for STD"
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
               Left            =   -405
               TabIndex        =   101
               Top             =   3000
               Visible         =   0   'False
               Width           =   1350
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
               Left            =   -660
               TabIndex        =   99
               Top             =   4080
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Variance"
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
               Left            =   0
               TabIndex        =   97
               Top             =   3600
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.Label lbAcquisition 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Acquisition Weight"
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
               Left            =   4920
               TabIndex        =   95
               Top             =   3960
               Width           =   1695
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qty LEFT for STD"
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
               Index           =   10
               Left            =   5280
               TabIndex        =   93
               Top             =   3360
               Width           =   1305
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pipette"
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
               Left            =   13215
               TabIndex        =   91
               Top             =   1800
               Width           =   660
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Lot"
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
               Left            =   720
               TabIndex        =   89
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Location"
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
               Index           =   1
               Left            =   5160
               TabIndex        =   88
               Top             =   1080
               Width           =   795
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date Arrived"
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
               Index           =   2
               Left            =   9240
               TabIndex        =   87
               Top             =   1080
               Width           =   1155
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Supplier EXP ( Date )"
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
               Index           =   3
               Left            =   12720
               TabIndex        =   86
               Top             =   1080
               Width           =   1785
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Purity"
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
               Left            =   1320
               TabIndex        =   85
               Top             =   1440
               Width           =   480
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Density"
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
               Left            =   9720
               TabIndex        =   84
               Top             =   1440
               Width           =   645
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Open Date"
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
               Left            =   840
               TabIndex        =   83
               Top             =   1800
               Width           =   990
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "MR EXP"
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
               Left            =   5295
               TabIndex        =   82
               Top             =   1800
               Width           =   615
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Status"
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
               Left            =   9840
               TabIndex        =   81
               Top             =   1800
               Width           =   555
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
            Left            =   1080
            TabIndex        =   118
            Top             =   12480
            Visible         =   0   'False
            Width           =   17175
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   22
               Left            =   7920
               TabIndex        =   197
               Top             =   5880
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Close Bottle"
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
                  TabIndex        =   198
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frBottles 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   1335
               Left            =   6120
               TabIndex        =   185
               Top             =   1920
               Visible         =   0   'False
               Width           =   5055
               Begin VB.Label lbWharehouse 
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
                  Left            =   1920
                  TabIndex        =   186
                  Top             =   530
                  Width           =   1155
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   20
               Left            =   0
               TabIndex        =   183
               Top             =   5880
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "All MR Bottle"
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
                  TabIndex        =   184
                  Top             =   120
                  Width           =   3015
               End
            End
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
               TabIndex        =   123
               Top             =   0
               Width           =   17175
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Select one bottle and press Add Acquisition"
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
                  Left            =   13170
                  TabIndex        =   125
                  Top             =   120
                  Width           =   3915
               End
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "MR Bottles Warehouse Table"
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
                  Left            =   6915
                  TabIndex        =   124
                  Top             =   105
                  Width           =   3435
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
               TabIndex        =   121
               Top             =   5880
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Goto Preparation"
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
                  TabIndex        =   122
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
               Left            =   11040
               TabIndex        =   119
               Top             =   5880
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Select Bottle"
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
                  TabIndex        =   120
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin FlexCell.Grid Grid3 
               Height          =   4935
               Left            =   0
               TabIndex        =   128
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
            TabIndex        =   117
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
            Left            =   13920
            TabIndex        =   113
            Top             =   240
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
               TabIndex        =   114
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
               TabIndex        =   111
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
                  TabIndex        =   112
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
               TabIndex        =   109
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
                  TabIndex        =   110
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
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   10095
            Index           =   0
            Left            =   960
            TabIndex        =   17
            Top             =   480
            Width           =   17175
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
               Index           =   16
               Left            =   3030
               Locked          =   -1  'True
               TabIndex        =   218
               Top             =   1680
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
               Index           =   15
               Left            =   7920
               Locked          =   -1  'True
               TabIndex        =   216
               Top             =   1680
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
               Index           =   14
               Left            =   3030
               Locked          =   -1  'True
               TabIndex        =   214
               Top             =   1320
               Width           =   2535
            End
            Begin VB.Frame frMS 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   1095
               Left            =   11040
               TabIndex        =   188
               Top             =   2280
               Visible         =   0   'False
               Width           =   6135
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
                  Height          =   640
                  Index           =   21
                  Left            =   3720
                  TabIndex        =   195
                  Top             =   300
                  Width           =   1695
                  Begin VB.Label lbCommandInside 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "Mother Soluiton"
                     BeginProperty Font 
                        Name            =   "Century Gothic"
                        Size            =   9
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C0C0C0&
                     Height          =   255
                     Index           =   21
                     Left            =   0
                     TabIndex        =   196
                     Top             =   180
                     Width           =   1695
                  End
               End
               Begin VB.TextBox txFormulation 
                  Alignment       =   2  'Center
                  BackColor       =   &H00F0F0F0&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00644603&
                  Height          =   300
                  Index           =   13
                  Left            =   1560
                  Locked          =   -1  'True
                  TabIndex        =   192
                  Top             =   660
                  Width           =   1575
               End
               Begin VB.TextBox txFormulation 
                  Alignment       =   2  'Center
                  BackColor       =   &H00F0F0F0&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00644603&
                  Height          =   300
                  Index           =   12
                  Left            =   1560
                  Locked          =   -1  'True
                  TabIndex        =   189
                  Top             =   300
                  Width           =   1575
               End
               Begin VB.Image IconMS 
                  Height          =   480
                  Left            =   5640
                  OLEDropMode     =   1  'Manual
                  Picture         =   "Preparation.frx":71D8
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   480
               End
               Begin VB.Line Line4 
                  BorderColor     =   &H00644603&
                  X1              =   5640
                  X2              =   240
                  Y1              =   160
                  Y2              =   160
               End
               Begin VB.Label lbFormulation 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00644603&
                  Caption         =   "  MR Qty  "
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   260
                  Index           =   13
                  Left            =   660
                  TabIndex        =   194
                  Top             =   660
                  Width           =   795
               End
               Begin VB.Label lbMRStock 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "mL"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00644603&
                  Height          =   240
                  Left            =   3240
                  TabIndex        =   193
                  Top             =   660
                  Width           =   270
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "mL"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00644603&
                  Height          =   240
                  Left            =   3240
                  TabIndex        =   191
                  Top             =   300
                  Width           =   270
               End
               Begin VB.Label lbFormulation 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00644603&
                  Caption         =   "  MS Volume  "
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   260
                  Index           =   12
                  Left            =   285
                  TabIndex        =   190
                  Top             =   300
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
               Index           =   6
               Left            =   11040
               TabIndex        =   179
               Top             =   8640
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "MR Calssification"
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
                  TabIndex        =   180
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
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00644603&
               Height          =   300
               Index           =   11
               Left            =   7920
               Locked          =   -1  'True
               TabIndex        =   176
               Top             =   3000
               Width           =   1575
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
               Index           =   10
               Left            =   3030
               TabIndex        =   173
               Top             =   2400
               Width           =   7455
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
               Index           =   9
               Left            =   7920
               Locked          =   -1  'True
               TabIndex        =   171
               Top             =   2040
               Width           =   2535
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
               TabIndex        =   159
               Top             =   8640
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
                  TabIndex        =   160
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
               Left            =   13200
               TabIndex        =   132
               Top             =   720
               Width           =   2535
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
               TabIndex        =   130
               Top             =   8640
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Image Image 
                  Height          =   480
                  Left            =   120
                  MousePointer    =   99  'Custom
                  OLEDropMode     =   1  'Manual
                  Picture         =   "Preparation.frx":A5BA
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
                  TabIndex        =   131
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frCommandInside 
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
               Index           =   14
               Left            =   11040
               TabIndex        =   126
               Top             =   8040
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Close Preparation"
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
                  TabIndex        =   127
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
               TabIndex        =   64
               Top             =   8040
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
               Index           =   5
               Left            =   3120
               TabIndex        =   62
               Top             =   8040
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
                  Index           =   5
                  Left            =   0
                  TabIndex        =   63
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
               Left            =   13200
               TabIndex        =   58
               Top             =   2040
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
               Left            =   13200
               Locked          =   -1  'True
               TabIndex        =   56
               Top             =   1320
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
               Index           =   0
               Left            =   7920
               Locked          =   -1  'True
               TabIndex        =   49
               Top             =   720
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
               Index           =   1
               Left            =   3000
               TabIndex        =   48
               Top             =   720
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
               Index           =   2
               Left            =   11760
               Locked          =   -1  'True
               TabIndex        =   47
               Top             =   1680
               Width           =   3975
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
               Left            =   7920
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   1320
               Width           =   2535
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00644603&
               Height          =   345
               Index           =   4
               Left            =   3030
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   3000
               Width           =   1335
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
               Left            =   840
               TabIndex        =   43
               Top             =   0
               Width           =   15255
               Begin VB.Line Line8 
                  BorderColor     =   &H00B0B0B0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   480
                  Y2              =   480
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hanna code parameters"
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
                  Width           =   2700
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
               Left            =   3030
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   2040
               Width           =   2535
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
               Top             =   8640
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Delete STD"
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
               Top             =   8640
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Add STD"
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
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   1335
               Left            =   6000
               TabIndex        =   23
               Top             =   5040
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
                  Top             =   530
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
               Top             =   8040
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
            Begin FlexCell.Grid Grid1 
               Height          =   3255
               Left            =   0
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   4440
               Width           =   17175
               _ExtentX        =   30295
               _ExtentY        =   5741
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
               Top             =   3840
               Width           =   17175
               Begin VB.Label lbRecipeDensity 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Select STD"
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
                  TabIndex        =   61
                  Top             =   120
                  Width           =   885
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
                  Caption         =   "Standards List Table"
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
                  Top             =   120
                  Width           =   17085
               End
            End
            Begin VB.Label lbSTDNote 
               Alignment       =   2  'Center
               BackColor       =   &H00644603&
               Caption         =   "Label13"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   0
               TabIndex        =   229
               Top             =   3600
               Visible         =   0   'False
               Width           =   17175
            End
            Begin VB.Line Line6 
               BorderColor     =   &H00B0B0B0&
               X1              =   16080
               X2              =   840
               Y1              =   1200
               Y2              =   1200
            End
            Begin VB.Label lbFormulation 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "MR Parameter"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00105010&
               Height          =   240
               Index           =   16
               Left            =   1650
               TabIndex        =   219
               Top             =   1680
               Width           =   1230
            End
            Begin VB.Label lbFormulation 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "FW MR Parameter"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00105010&
               Height          =   240
               Index           =   15
               Left            =   6315
               TabIndex        =   217
               Top             =   1680
               Width           =   1530
            End
            Begin VB.Label lbFormulation 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hanna Formula"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00105010&
               Height          =   240
               Index           =   14
               Left            =   1560
               TabIndex        =   215
               Top             =   1320
               Width           =   1320
            End
            Begin VB.Label lbFormulation 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00644603&
               Caption         =   "  Tot MR  "
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   11
               Left            =   6615
               TabIndex        =   178
               Top             =   3000
               Width           =   1185
            End
            Begin VB.Label lbMRUnit 
               BackStyle       =   0  'Transparent
               Caption         =   "mL"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00644603&
               Height          =   255
               Left            =   9600
               TabIndex        =   177
               Top             =   3000
               Width           =   1095
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "mL"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00644603&
               Height          =   255
               Left            =   4440
               TabIndex        =   175
               Top             =   3000
               Width           =   615
            End
            Begin VB.Image Image2 
               Height          =   360
               Left            =   5040
               Picture         =   "Preparation.frx":D99C
               Top             =   3000
               Width           =   360
            End
            Begin VB.Label lbFormulation 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Preparation Note"
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
               Left            =   900
               TabIndex        =   174
               Tag             =   "Note"
               Top             =   2400
               Width           =   1980
            End
            Begin VB.Label lbFormulation 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Storage STD"
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
               Left            =   6255
               TabIndex        =   172
               Top             =   2040
               Width           =   1590
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
               Height          =   255
               Index           =   8
               Left            =   11520
               TabIndex        =   133
               Top             =   720
               Width           =   1620
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
               Left            =   12240
               TabIndex        =   59
               Top             =   2040
               Width           =   825
            End
            Begin VB.Label lbFormulation 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measurement Unit"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00105010&
               Height          =   240
               Index           =   6
               Left            =   11535
               TabIndex        =   57
               Top             =   1320
               Width           =   1560
            End
            Begin VB.Label lbFormulation 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Preparation Time"
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
               Index           =   0
               Left            =   6255
               TabIndex        =   55
               Top             =   720
               Width           =   1590
            End
            Begin VB.Label lbFormulation 
               Alignment       =   1  'Right Justify
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
               Left            =   1620
               TabIndex        =   54
               Top             =   720
               Width           =   1260
            End
            Begin VB.Label lbFormulation 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "STD MATRIX"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00105010&
               Height          =   240
               Index           =   2
               Left            =   10665
               TabIndex        =   53
               Top             =   1680
               Width           =   1035
            End
            Begin VB.Label lbFormulation 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "FW Hanna Parameter"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00105010&
               Height          =   240
               Index           =   3
               Left            =   6015
               TabIndex        =   52
               Top             =   1320
               Width           =   1830
            End
            Begin VB.Label lbFormulation 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00644603&
               Caption         =   "  STD Volume  "
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   4
               Left            =   915
               TabIndex        =   51
               Top             =   3000
               Width           =   1905
            End
            Begin VB.Label lbFormulation 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "STD Exp ( Days )"
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
               Left            =   1620
               TabIndex        =   50
               Top             =   2040
               Width           =   1260
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00D0D0D0&
               X1              =   120
               X2              =   17280
               Y1              =   7800
               Y2              =   7800
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Procedure : Standard Preparation"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   18975
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
      Top             =   10920
      Width           =   19215
      Begin VB.Timer TimerBeginForm 
         Interval        =   200
         Left            =   8400
         Top             =   120
      End
      Begin VB.Label lbClosed 
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
         Left            =   120
         TabIndex        =   187
         Top             =   120
         Visible         =   0   'False
         Width           =   5310
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
         Picture         =   "Preparation.frx":F40E
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   3
         Left            =   15600
         MousePointer    =   99  'Custom
         Picture         =   "Preparation.frx":127F0
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   4
         Left            =   18240
         MousePointer    =   99  'Custom
         Picture         =   "Preparation.frx":15BD2
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
      Begin ChemicalMR.ucScrollAdd ucScrollAdd1 
         Left            =   2760
         Top             =   480
         _ExtentX        =   1138
         _ExtentY        =   423
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
            Picture         =   "Preparation.frx":18FB4
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
      Begin VB.Label lbHannaCode 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preparation"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   870
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Visible         =   0   'False
         Width           =   19170
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
         Left            =   13620
         TabIndex        =   3
         Top             =   195
         Width           =   5310
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

Private IDBottle As Long

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
Private SelectedHannaCode As String





Private lRowBottle As Long
Private lColBottle As Long

Private lRowRecipe As Long
Private lColRecipe As Long


Private lRowCombo As Long

Private IndexRecipe As Integer
Private indexMix As Integer
Private IndexComponent   As Integer



Private uRecipe As RecipeType
Private uPreparation As RecipeForProduction
Private uHannaCode As HannaCode



Private SettingName As String
Private bImportata As Boolean
Private bIfDataPath As Boolean
Private bfrInsideMoveTop As Boolean

Private bCancelUpdate As Boolean


Private STDNumber As String

Private Frame3Top As Long
Private Grid1Height As Long
Private PreparationID As Long

Private userAcquisition As PrepAcquisition
Private userAcquisitionClean As PrepAcquisition
Private uBottle() As WareHouseEntry

Private AcquisitionID As Long
Private AcquisitionSTDNumber As String
Private AcquisitionWeight As String
Private lAcquisitionRow As Long


Private UserSTD As String
Private UserSTDValue As Double
Private lSTDRow As Long
Private lSTDCol As Long
Private bPreparationClosed As Boolean
Private bManualClosepreparation As Boolean

Private NotesID As Long

Private lMSRow As Long

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


Public Function DoShow(ByVal HannaCode As String, ByVal FileName As String, Optional ByVal Preparation_ID As Long, Optional bClosed As Boolean, Optional bDoClosePreparation As Boolean) As Boolean

    On Error GoTo ERR_SHOW
    
    
    bPreparationClosed = bClosed
    bManualClosepreparation = bDoClosePreparation
    m_rc = False
    mOk
    PreparationID = Preparation_ID
  
    SettingName = FileName
    bImportata = IIf(FileName <> "", True, False)

    SelectedHannaCode = HannaCode
    SelectedCode = GetMRCodeFromHannaCode(HannaCode)


    lbHannaCode = SelectedHannaCode & " | " & SelectedCode
    lbHannaCode.Visible = True
    
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

Private Sub Grid1_Click()
Dim strSTD As String
With Grid1
    If STDNumber = "" Then Exit Sub
    If (lSTDCol < 3 And lSTDCol > 0) Or lSTDCol = 9 Then
        If STDNumber <> 9999 And UserSTDValue <> 9999 Then
            strSTD = .Cell(lSTDRow, lSTDCol).Text
            If F_InputBox.DoShow("Enter Value:", "STD : " & .Cell(0, lSTDCol).Text, , , , strSTD, , True) Then
            
            
                If lSTDRow > UBound(uRecipe.STD) Then
                        
                        ReDim Preserve uRecipe.STD(lSTDRow)
                        
                End If
                .Cell(lSTDRow, lSTDCol).Text = strSTD
            
                Select Case lSTDCol
                    
                    Case 1
                        uRecipe.STD(lSTDRow).NUMBER = Int(strSTD)
                    Case 2
                        uRecipe.STD(lSTDRow).Value = CDbl(strSTD)
                        Image2_Click
                End Select
            
            End If
        
        
        End If
    End If
End With


End Sub



Private Sub Grid1_DblClick()

AddAcquisition

End Sub

Private Sub Grid1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

lSTDRow = 0
lSTDCol = 0
STDNumber = 9999
UserSTDValue = 9999

    frCommandInside(5).Visible = False
    frCommandInside(14).Visible = False
    
    
    If FirstRow > 0 Then
            lSTDRow = FirstRow
            STDNumber = Trim(Grid1.Cell(FirstRow, 1).Text)
            UserSTDValue = IIf(IsNull(Trim(Grid1.Cell(FirstRow, 2).Text)) Or Trim(Grid1.Cell(FirstRow, 2).Text) = "", 9999, Trim(Grid1.Cell(FirstRow, 2).Text))
            frCommandInside(5).Visible = IIf(STDNumber <> "", True, False)
            lSTDCol = FirstCol
    End If
    
    With Grid1
        Dim i As Integer
        For i = 1 To .Rows - 1
        
          
            .Cell(i, 1).BackColor = &HE0E0E0
            .Cell(i, 2).BackColor = &HE0E0E0
            .Cell(i, 3).BackColor = &HE0E0E0
            
            .Cell(i, 1).ForeColor = vbBlack
            .Cell(i, 2).ForeColor = vbBlack
            .Cell(i, 3).ForeColor = vbBlack
            
            
            If i = FirstRow Then
                
                .Cell(i, 1).BackColor = &H886010
                .Cell(i, 2).BackColor = &H886010
                .Cell(i, 3).BackColor = &H886010
                
                .Cell(i, 1).ForeColor = vbWhite
                .Cell(i, 2).ForeColor = vbWhite
                .Cell(i, 3).ForeColor = vbWhite
                       
            End If
            
        Next
    
    End With
End Sub



Private Sub Grid2_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
On Error GoTo ERR_GRID
AcquisitionID = 0
AcquisitionSTDNumber = ""
AcquisitionWeight = 0
lAcquisitionRow = 0
If FirstRow > 0 Then
    lAcquisitionRow = FirstRow
    AcquisitionID = Grid2.Cell(FirstRow, 11).Text
    AcquisitionSTDNumber = Grid2.Cell(FirstRow, 4).Text
    AcquisitionWeight = CDbl(Grid2.Cell(FirstRow, 7).Text)
    frCommandInside(11).Visible = Not (bPreparationClosed)
    
End If
ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_GRID:
    PopupMessage 2, Err.Description
    Exit Sub

End Sub



Private Sub Grid3_DblClick()

Dim ID As Long

 With Grid3
    If lRowBottle > 0 Then
        
        If F_MsgBox.DoShow("Select bottle?", .Cell(lRowBottle, 2).Text) Then
        
            ID = Grid3.Cell(lRowBottle, 18).Text
             
            
            ReDim uBottle(0)
             
                 
            Select Case uPreparation.MsType
                
                Case 0
                     SelectBotteForAcquisition (ID)
                Case 1
                    ' MS1
                    Call GetDatabaseWareHouseEntry(ID, 0, True, uBottle)
                    Call CreateMotherSolution(uBottle(0))
                Case 2
                    ' MS2
                    Call GetDatabaseWareHouseEntry(ID, 0, True, uBottle)
                    Call CreateMotherSolution(uBottle(0))
            
            End Select
             
             
             
             
          
        End If

    End If
End With
End Sub


Private Sub SelectBotteForAcquisition(ByVal ID As Long)


    If AddBottleInAcquisition(ID) Then
       
       ' back to acquisition....
       frCommandInside_Click 12
       ' acquisisco dalla bottiglia...
       DoEvents
       
       frCommandInside(13).Visible = False
       Call AcquireMRfromBottle(uBottle(0))
       
    End If
        
        
End Sub

Private Sub Grid3_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
'-----------------
' bottles grid
'-----------------
IDBottle = 0
lRowBottle = FirstRow
lColBottle = FirstCol

lbCommandInside(13).Visible = False
frCommandInside(15).Visible = False
frCommandInside(22).Visible = False

If FirstRow > 0 Then

lbCommandInside(13).Visible = True
frCommandInside(15).Visible = True
frCommandInside(22).Visible = True
   IDBottle = Grid3.Cell(lRowBottle, 18).Text

End If

End Sub



Private Sub Grid6_DblClick()

frCommandInside_Click 27

End Sub

Private Sub Image_Click()
frExcel_Click
End Sub

Private Sub Image2_Click()

RicaricoGrid1

End Sub

Private Sub lbChemical_Change()
lbInside(2) = lbChemical
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
    
    If MyOperatore.Name <> "" Then
    Else
        
    End If
   
    uPreparation.Operator = MyOperatore.Name
    
   
 
    
    If bImportata Then
    
        GetFileInfo
        
    Else
    
        
        Call SetInitPreparation
        Call GetHannaCodeDataFromDatabase
   
   
       'SetSettingName
       
    End If
    
    
    '-----------------------------------------------------
    '    chiudi la preparatione da Database Preparation
    '-----------------------------------------------------
    If bManualClosepreparation Then
    
            frCommandInside(14).Visible = True
            
            frCommandInside_Click 14
    
    End If


End Sub

Private Sub GetHannaCodeDataFromDatabase()

On Error GoTo ERR_GET:


uPreparation.QtyToProduce = IIf(uPreparation.HannaCode.STDVolume = "", 500, uPreparation.HannaCode.STDVolume)

With uPreparation.HannaCode

        txFormulation(2) = .STDMatrix
        txFormulation(3) = .FWHannaParameter
        txFormulation(4) = .STDVolume
        txFormulation(5) = .STDExp
        txFormulation(6) = .MeasurementUnit
        txFormulation(9) = .STDStorage
        
        txFormulation(14) = .Hannaformula
        txFormulation(15) = .MR.FWParameter
        txFormulation(16) = .MR.Parameter
          


End With

 With uPreparation
 
        txFormulation(0) = .HourPrep
        txFormulation(7) = .Operator
        txFormulation(10) = .Note
        txFormulation(1) = IIf(.DataPrep <> 0, FormatDataLAT(CStr(.DataPrep)), "")
        txFormulation(8) = .PrepWeek
        txFormulation(4) = .QtyToProduce
        
        If .MS.Volume = 0 Then
            .MS.Volume = 100
         
        End If
        If .MotherSol.QtyLeft = 0 Then
            .MotherSol.QtyLeft = .MS.Volume
            
        End If
        
        txFormulation(12) = .MotherSol.QtyLeft
        
        '''''.QtyProduced = txFormulation(4)
    End With

ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_GET:
    MsgBox Err.Description
    Resume Next

End Sub

Private Sub SetHannaCodeDataPreparation()


    With uPreparation.HannaCode
    

        .STDVolume = txFormulation(4)
        .STDExp = txFormulation(5)
        .STDStorage = txFormulation(9)
        
      
        
    End With
    
    With uPreparation
        
        .Operator = txFormulation(7)
        .Note = txFormulation(10)
        .HourPrep = txFormulation(0)
        .DataPrep = FormatDataLAT(txFormulation(1))
        .PrepWeek = txFormulation(8)
        .QtyToProduce = txFormulation(4)
        '''''.QtyProduced = txFormulation(4)
        .FileName = SettingName
    
    End With


End Sub


Private Sub InitForm()




  
    uPreparation = uPreparationClean
    
   
    lRowBottle = 0
    lColBottle = 0
    lRowRecipe = 0
    lColRecipe = 0
   
    lRowCombo = 0
    IndexRecipe = 0
    indexMix = 0
    
   
    
    
    Dim Grid(10) As Grid
    
    Set Grid(0) = Grid1
    Set Grid(1) = Grid2
    Set Grid(2) = Grid3
    'Set Grid(3) = Grid4
    Set Grid(4) = Grid5
    Set Grid(5) = Grid6
  
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



Private Sub SetInitPreparation()
'---------------------------------
' preparation senza file.....
'---------------------------------

  With uPreparation
 
    Call SetHannaCodeByCode(SelectedHannaCode, .HannaCode)

    .MRCode = .HannaCode.MR.Code
    
    lbSTDNote = .HannaCode.STDNote
    lbSTDNote.Visible = IIf(Len(Trim(.HannaCode.STDNote)) > 0, True, False)
    lbSTDNote_AQ = lbSTDNote
    lbSTDNote_AQ.Visible = lbSTDNote.Visible

   
   ' inizializzo la "ricetta" con il counter del DB HannaCode
 
    .Recipe.STDcount = .HannaCode.STDcount
    
    ReDim .Recipe.STD(.Recipe.STDcount)
    
    .Recipe.STD = .HannaCode.STD
    .Recipe.STDUnit = .HannaCode.MeasurementUnit
    Dim i As Integer

    Call SetInitSTDTheoreticalWeight(uPreparation)
    Call GetPreparationSTDGrid(Grid1, uPreparation)
    
    
        uRecipe = .Recipe
        uHannaCode = .HannaCode
        
        uRecipe.Code = uPreparation.MRCode
      
            
        Call SetFormMRMS
   End With
    
End Sub


Private Sub SetFormMRMS()
            
    With uPreparation
        
        
       Call STDRefreshTotalMSVolume
       
        
        lbFormulation(11) = "  Tot MR  "

        Select Case .MsType
            
            Case 0
                frMS.Visible = False
            Case 1
                frMS.Visible = True
                lbFormulation(11) = "  Tot MS  "
            Case 2
                frMS.Visible = True
                lbFormulation(11) = "  Tot MS  "
            
        
        End Select

        
    lbHannaCode.Visible = True
    lbHannaCode = .HannaCode.Code & " | " & .HannaCode.MR.Code
    
    
    End With
    
End Sub



Private Sub Form_Load()

    PBContainer.Top = 0
    PBContainer.Left = 0
    PBContainer.Width = Me.Width
  
    
    

    ucScrollAdd1.AddScroll PBContainerViewport(0)
    ucScrollAdd1.TrackMouseWheel Vertical
    ucScrollAdd1.ResizeTargetOnFormResize 0, 0
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
 

    PBContainerViewport(0).ZOrder
    PBFooter.ZOrder
    
    
    
    
    
End Sub

Private Sub Form_Resize()




    
    lbWait.Left = Me.Width / 2 - lbWait.Width / 2
    PBTitle.Width = Me.Width
    PBFooter.Top = Me.ScaleHeight - PBFooter.Height
    PBFooter.Width = Me.Width
 
    frInside(6).Left = frInside(5).Left
    frInside(6).Top = frInside(5).Top
    'Resize the container (needed to show the full bottom box on maximized state)
    'First resize our container
    ucScrollAdd1.ContainerW = Me.ScaleWidth
    'But also need to resize PBContainer wich hide the width of the bottom box

    
    
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
        Case 1, 0, 8
            txFormulation(Index).BackColor = IIf(rc, vbWhite, vbRed)
            
        Case 12
            uPreparation.HannaCode.MS1vol = txFormulation(12)
            uPreparation.HannaCode.MS2vol = txFormulation(12)
    End Select



End Sub

Private Sub txFormulation_Click(Index As Integer)
Dim Answer As String
Dim Selected As String
Dim sString As String
Dim bNumber As Boolean
Dim OldValue As String

Selected = "Preparation"
Answer = txFormulation(Index)
sString = lbFormulation(Index)

bNumber = IIf(Index = 4, True, False)

If Index = 14 Then Exit Sub
If Index = 15 Then Exit Sub
If Index = 16 Then Exit Sub
If Index = 13 Then Exit Sub
If Index = 11 Then Exit Sub
If Index = 3 Then Exit Sub
If Index = 6 Then Exit Sub
If Index = 2 Then Exit Sub


If bPreparationClosed Then Exit Sub

If Index = 4 And uPreparation.Recipe.AcquisitionCount > 0 Then

    PopupMessage 2, "Preparation already started. Cannot change Qty. Delete preparation and start again..."
    Exit Sub
    
End If

If Index = 0 Then If Answer = "" Then Answer = FormatDateTime(Now(), vbShortTime)
If Index = 1 Then If Answer = "" Then Answer = FormatDataLAT(Now())
If Index = 8 Then If Answer = "" Then Answer = PreparationWeek(Now())


If Index = 7 Then
    
    If frmLogin.DoShow Then
        txFormulation(7) = MyOperatore.Name
        Exit Sub
    Else
    
        Exit Sub
    End If
    
End If

        OldValue = Answer
        
        If Index = 4 Then
            ' se voglio cambiare il volume ma ho già iniziato?
            If CheckPreparation Then
            Else
                If F_MsgBox.DoShow("Preparation already started" & vbCrLf & "change volume and reset preparation?") Then
                Else
                    Exit Sub
                End If
            End If
        
        End If
        
        
        
        If F_InputBox.DoShow(sString, Selected, , , , Answer, , bNumber, Me.Top) Then
        
            txFormulation(Index) = Answer
            
            Select Case Index
                Case 0
                   
                    uPreparation.HourPrep = FormatDateTime(Answer, vbShortTime)
                    txFormulation(0) = uPreparation.HourPrep
                    
                Case 1
                    ' isdate?
                    If IsDate(Answer) Then
                        txFormulation(Index) = FormatDataLAT(Answer)
                        uPreparation.DataPrep = txFormulation(Index)
                        uPreparation.PrepWeek = PreparationWeek(uPreparation.DataPrep)
                        txFormulation(8) = uPreparation.PrepWeek
                        
                        If uPreparation.HourPrep = "" Then
                            uPreparation.HourPrep = FormatDateTime(Now(), vbShortTime)
                            txFormulation(0) = uPreparation.HourPrep
                            
                        End If
                    Else
                    
                        PopupMessage 2, "Please enter a valid Date...", , True
                    End If
                Case 4
                    ' cambia il volume!!!!!!!
                    If Answer > 0 Then
                    
                            uPreparation.QtyToProduce = Answer
                            txFormulation(4) = uPreparation.QtyToProduce
                    
                        Else
                            PopupMessage 2, "Please enter a valid Volume...", , True
                    End If
                    
                    Image2_Click
                    
                Case 8
                    uPreparation.PrepWeek = Answer
                     txFormulation(8) = uPreparation.PrepWeek
                     
                Case 12
                    If OldValue = Answer Then
                    Else
                        
                        uPreparation.MotherSol = MotherSolutionClean
                        Call SetFrameMotherSolution("")
                        Image2_Click
                        
                    End If
            End Select
        End If
        
        




End Sub

Private Function CheckPreparation() As Boolean
Dim rc As Boolean
With Grid1
    rc = True
    If .Rows > 0 Then
        rc = IIf(.Cell(.Rows - 1, 4).Text > 0, False, True)
    End If

End With
CheckPreparation = rc

End Function



Private Sub txFormulation_LostFocus(Index As Integer)
    
    Select Case Index
        
        Case 1, 2, 3
           
    
    End Select
    

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
    
    
    If ucScrollAdd1.UCScrollV.Value <= frInside(0).Top Then
        IndexVisibleFrame = 0
    ElseIf ucScrollAdd1.UCScrollV.Value > frInside(0).Top And ucScrollAdd1.UCScrollV.Value <= frInside(1).Top Then
        IndexVisibleFrame = 1
    
    ElseIf ucScrollAdd1.UCScrollV.Value > frInside(1).Top And ucScrollAdd1.UCScrollV.Value <= frInside(2).Top Then
        IndexVisibleFrame = 2
    ElseIf ucScrollAdd1.UCScrollV.Value > frInside(2).Top And ucScrollAdd1.UCScrollV.Value <= frInside(3).Top Then
        IndexVisibleFrame = 3
   ' ElseIf ucScrollAdd1.UCScrollV.Value > frInside(3).Top And ucScrollAdd1.UCScrollV.Value <= frInside(4).Top Then
       ' IndexVisibleFrame = 4
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


Private Sub Form_Activate()
Me.WindowState = MainWindowState
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
        
    Case 3
        ' Previous
         If IndexVisibleFrame > 1 Then
            MyIndex = IndexVisibleFrame - 1
            
            
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(MyIndex).Top - 680
        Else
            ucScrollAdd1.UCScrollV.ScrollToValue 0
         End If
    
    
    
    Case 4
        ' forward
        If IndexVisibleFrame < frInside.UBound Then
            MyIndex = IndexVisibleFrame + 1
            
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(MyIndex).Top - 680
        Else
            ucScrollAdd1.UCScrollV.ScrollToValue 0
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
Dim rc As Boolean
Dim UserSTDNumber As String

    txQRCode.Visible = False


    bCancelUpdate = False
        
    Select Case Index
        Case 0
            'add STD
            Call AddSTD
        Case 1
            ' delete component
            Call DeleteSTD
        Case 2
            ' scroll top
            
            ucScrollAdd1.UCScrollV.ScrollToValue 0
        Case 3
            ' scroll to acquisitions
            frInside(1).ZOrder
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(1).Top - 680
        Case 4
            Call SavePreparation
        Case 5
            ' add acquisition
            Call AddAcquisition
        Case 6
            ' open product classification
            Call OpenProductClassification(uRecipe.Code, 1)
        Case 7
            ' exit acquisition
            Call AcquireAcquisition(False)
            
        Case 8
            ' save acquisition
             Call AcquireAcquisition(True)
        Case 9
        
        
            ' pipette
            Call GotoPipetteGrid
            
            
        Case 10
            ' add acquisition
            frCommandInside_Click 5
        Case 11
            ' delete acquisition
            Call DeleteAcquisition
        Case 12
            ' back to acquisition
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(2).Top - 680
            frCommandInside(12).Visible = False
        Case 13
            ' SCANN BARCODE
            Call ScanQRCode
        Case 14
            ' ClosePreparation
            
            Call ClosePreparation
        Case 15
            Grid3_DblClick
            
        Case 16
            ' goto preparation...
             ucScrollAdd1.UCScrollV.ScrollToValue 0
            frInside(3).Visible = False
        Case 17
            Call ClearRevisionForm
        Case 18
            If SettingName <> "" Then
                AddcmbRevType
                lbInside(3).ForeColor = vbWhite
                Call GetPreparationNotes(Grid4, SettingName)
                 frExcel2.Visible = IIf(Grid4.Rows > 1, True, False)
                Frame6.Visible = IIf(Grid4.Rows > 1, False, True)
            
                Call ClearRevisionForm
                frInside(4).Visible = True
                ucScrollAdd1.UCScrollV.ScrollToValue frInside(4).Top - 680
            Else
                PopupMessage 2, "Please save Preparation first..."
            End If
            
            
        Case 19
           ' goto BOTTLE table
            
            Call AcquisitionBottle
            
        Case 20
            ' grid3 ALL BOTTLES
            Call SetBottleTable(True, False)
            
        Case 21
         ' goto BOTTLE table MS
         
            Call SetMotherSolution
         
          
            
            
        Case 22
            Call CloseBottleInTable
            
            
        Case 25
            ' scroll to STDs
            frInside(5).Visible = False
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(0).Top - 680
            
        Case 27
              ' scroll to acquisition
            frInside(6).Visible = False
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(2).Top - 680
    End Select
End Sub




Private Sub frInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


Dim i As Integer
    For i = 0 To frCommandInside.UBound

            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 4 Or i = 8 Or i = 14 Then
                frCommandInside(i).BackColor = &H8000&
            ElseIf i = 13 Then
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
            If i = 4 Or i = 8 Or i = 14 Then
            
                frCommandInside(i).BackColor = &H20A020
            ElseIf i = 13 Then
            End If
            
        Else
            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 4 Or i = 8 Or i = 14 Then
                frCommandInside(i).BackColor = &H8000&
            ElseIf i = 13 Then
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
        PicMenu(i).BackColor = &H644603
    End If
Next

End Sub


'-----------------------------------------------------------------------------------------------
'   SettingName
'-----------------------------------------------------------------------------------------------

Private Sub SetSettingName()
'
If txFormulation(1) <> "" And txFormulation(8) <> "" And txFormulation(0) <> "" Then
    SettingName = FormatNomeFile(Trim(uPreparation.HannaCode.Code) & "." & Trim(uPreparation.MRCode) & "." & txFormulation(1) & "." & txFormulation(0) & "." & txFormulation(8)) & "." & USER_ESTENSIONE_PREPARATION
End If

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
Dim bAcquisitiTutti As Boolean

On Error GoTo ERR_GET:
   
rc = True



        
    lbWait.Visible = True
    
    Debug.Print USER_PATH
    
    uPreparation = uPreparationClean
    
    
    If SettingName = "" Then
            MessageInfoTime = 2000
            PopupMessage 2, "Warning : File non found! ", , True, "Preparation"
            rc = False
            GoTo ERR_END
    
    End If
    
    If FileExists(USER_TEMP_PATH & SettingName) Then
        USER_PATH = USER_TEMP_PATH
    ElseIf FileExists(USER_DATA_PATH & SettingName) Then
        USER_PATH = USER_DATA_PATH
    Else
        rc = False
        PopupMessage 2, "No file Preparation found...", , True, SettingName
        GoTo ERR_END
        
    End If
    
    
    If GetSettingData(SettingName, "iRecipeForProduction", "bClosed", False) Then


            
            blTable.Visible = False
            
            PicMenu(1).Visible = False

            bPreparationClosed = True
            
            
            
    End If
    
    
    
    Call SetClosedPreparation(bPreparationClosed)
    Call PreparationGetSetting(uPreparation, SettingName, SelectedHannaCode)
     
    
    With uPreparation
    
        PreparationID = .ID
       .Recipe.Code = .MRCode

        uRecipe = .Recipe
        
         
        

        GetHannaCodeDataFromDatabase
        
        If bPreparationClosed Then
            lbClosed = "Closed | " & FormatDataLAT(uPreparation.CloseDate)
            lbClosed.Visible = True
        Else
            lbClosed = "Open | " & FormatDataLAT(CStr(uPreparation.DataPrep))
            lbClosed.Visible = True
        End If
        
        
        
        
       Select Case .MsType
            
            Case 0
                frMS.Visible = False
            Case 1
                frMS.Visible = True
            Case 2
                frMS.Visible = True
                
        
        End Select
 
        
      '-----------------------
      ' mother solution
      '------------------------
      Call SetFrameMotherSolution(.MotherSol.Code)
      
      
      
        lbSTDNote = .HannaCode.STDNote
        lbSTDNote.Visible = IIf(Len(Trim(.HannaCode.STDNote)) > 0, True, False)
        lbSTDNote_AQ = lbSTDNote
        lbSTDNote_AQ.Visible = lbSTDNote.Visible

    End With
    

    
    
    
    Call FillGridPreparationFromFile(Grid2, uPreparation, 2, PreparationID)
    Call FillGridPreparationFromFile(Grid1, uPreparation, 1, PreparationID, bAcquisitiTutti)
    
    frCommandInside(14).Visible = bAcquisitiTutti
    
    
    uRecipe.TotalWeight = uPreparation.Recipe.TotalWeight
    uRecipe.ActualWeight = uPreparation.Recipe.ActualWeight
  
  
     Call SetFormMRMS
  

    blTable.Visible = True

    If uRecipe.bUmMassa Then

    End If
    
    
 
    
   
ERR_END:

    On Error GoTo 0
     uRecipe = uPreparation.Recipe
    lbWait.Visible = False
    GetReceiptFromFile = rc
    Exit Function
ERR_GET:
    rc = False
    Resume Next
    
End Function




'-----------------------------------------------------------------------------------------------
'
'
'                                   Acquisition

'
'-----------------------------------------------------------------------------------------------


Private Sub AcquisitionBottleMS()
   Call SetBottleTable(False, True)

   frInside(3).Left = frInside(1).Left
   frInside(3).Top = frInside(1).Top

   frInside(3).ZOrder
   frInside(3).Visible = True
   ucScrollAdd1.UCScrollV.ScrollToValue frInside(3).Top - 680
   
   
End Sub



Private Sub AcquisitionBottle()


   Call SetBottleTable(False, False)

   frInside(3).Left = frInside(1).Left
   frInside(3).Top = frInside(1).Top

   frInside(3).ZOrder
   frInside(3).Visible = True
   ucScrollAdd1.UCScrollV.ScrollToValue frInside(3).Top - 680
   
   
   
 
End Sub



Private Sub SetBottleTable(ByVal bTutti As Boolean, ByVal bMotherSolution As Boolean)

    frBottles.Visible = False

    lbCommandInside(13).Visible = False
    frCommandInside(15).Visible = False
    frCommandInside(22).Visible = False
    frCommandInside(20).Visible = False

    IDBottle = 0
    
    Dim strLott As String
    If bTutti Then
        strLott = ""
        lbWharehouse = "Empty list"
        frCommandInside(20).Visible = False
        
    Else
    
        If uRecipe.AcquisitionCount > 0 Then
            strLott = uRecipe.Acquisitions(1).MRLot
        End If
        
        lbWharehouse = "Bottle with Lot : " & strLott
        
        frCommandInside(20).Visible = True
        
        
    End If
    
    Call GetStockFromDatabase(Grid3, False, uRecipe.Code, strLott, bMotherSolution, uPreparation.MS.Qty, bTutti)
    
    If Grid3.Rows > 1 Then
    
    Else
        MessageInfoTime = 2000
        PopupMessage 2, "NO ETRIES! Check MR Stock Bottles for this Code", , , uRecipe.HannaCode.MR.Code
    
    End If
    
    frCommandInside(20).Visible = IIf(Grid3.Rows > 1, False, True)
    frBottles.Visible = IIf(Grid3.Rows > 1, False, True)

End Sub



Private Sub ClearAndFillAcquisition()

    Call ClearAcquisition
    Call fillAcquisitionDetail

End Sub


Private Function AddAcquisition()
 Dim STDType As Integer
 
 
     If STDNumber = "" Or bPreparationClosed Then Exit Function
     
    
    uBottle = MyWareHouseEntryCleanArray
    
    ReDim uBottle(0)
    
    
    
    
    With uRecipe.STD(lSTDRow)
    
        .TheoreticalWeight = Grid1.Cell(lSTDRow, 4).Text
        .RealWeight = Grid1.Cell(lSTDRow, 5).Text
       
        .Variance = .RealWeight - .TheoreticalWeight
     
        .Value = Grid1.Cell(lSTDRow, 2).Text
        
    
        If .Value > 0 And .Variance = 0 Then
            
            If .TheoreticalWeight = .RealWeight Then
                ' ho acquisito tutto....
                PopupMessage 2, "STD Acquired...", , , "STD " & .NUMBER
                Exit Function
            End If
      
        
        
        End If
    
    
    End With

    
    ClearAndFillAcquisition
    
    frInside(2).ZOrder
    frInside(2).Visible = True

    
    With uPreparation
        Select Case .MsType
            
            Case 0
              
            Case 1
                Call GetMSAcquisition
                Exit Function
                         
            Case 2
                Call GetMSAcquisition
                Exit Function
            
            
        
        End Select
    End With

    txAcquisition(0).SetFocus
    ucScrollAdd1.UCScrollV.ScrollToValue frInside(2).Top - 680

   
    If txAcquisition(10) = 0 And uRecipe.STD(lSTDRow).Value = 0 Then
        ' acquisisco immediatamente..
        ' ma cosa???
        frCommandInside(8).Visible = True
        If F_MsgBox.DoShow("STD Number : " & uRecipe.STD(lSTDRow).NUMBER & vbCrLf & "STD Value : " & uRecipe.STD(lSTDRow).Value & vbCrLf & "Acquire?", lbInside(2)) Then
        
                 txAcquisition(13) = 0
        
                Call AcquireAcquisition(True)
        
        End If
        
    End If
    



End Function

Private Sub fillAcquisitionDetail()

Dim STDValue As Double
With uRecipe.STD(lSTDRow)

    .RealWeight = Grid1.Cell(lSTDRow, 5).Text

    .ActualWeight = -Grid1.Cell(lSTDRow, 6).Text ' FormatNumber(.TheoreticalWeight - .RealWeight, 3)
    
    STDValue = .Value

End With



With Grid1
    txAcquisition(15) = .Cell(lSTDRow, 1).Text
    txAcquisition(16) = .Cell(lSTDRow, 2).Text
    txAcquisition(19) = .Cell(lSTDRow, 3).Text
    
    txAcquisition(10) = uRecipe.STD(lSTDRow).ActualWeight '  .Cell(lSTDRow, 3).Text
    
    txAcquisition(14) = .Cell(lSTDRow, 5).Text
    txAcquisition(11) = .Cell(lSTDRow, 6).Text
    txAcquisition(12) = .Cell(lSTDRow, 7).Text
End With


End Sub


Private Sub txStock_Change(Index As Integer)
Dim rc As Boolean
rc = IIf(Len(txStock(Index)) > 0, True, False)
Select Case Index
    Case 0
    
        frCommandInside(8).Visible = rc

End Select
End Sub

Private Sub txStock_Click(Index As Integer)

        txAcquisition_Click Index

End Sub

Private Sub txAcquisition_Click(Index As Integer)
Dim rc As Boolean
Dim strPipette As String





    rc = IIf(Len(txStock(1)) > 0 And Len(txAcquisition(0)) > 0, True, False)
    
    If Index = 9 Then
        '-------------
        ' pipette....
        '-------------
        strPipette = txAcquisition(9)
        
        If F_InputBox.DoShow("Enter Code :", "STD preparation : Pipette", , , , strPipette) Then
        
            txAcquisition(9) = strPipette
        
        End If
        
        
        Exit Sub
        
    
    End If
    
    If Index = 17 Then
        '-------------
        ' NOTE....
        '-------------
        strPipette = txAcquisition(17)
        
        If F_InputBox.DoShow("Enter Note :", "Acquisition Details", , , , strPipette) Then
        
            txAcquisition(17) = strPipette
        
        End If
        
        
        Exit Sub
        
    
    End If
    
    If uBottle(0).EntryBottle = "" Then
        
        PopupMessage 2, "Please select a bottle first...."
        
        Exit Sub
    
    End If
    
    If Index = 13 And rc Then
    
        ' acquisisco l'MR dalla bottiglia....
        Dim strValue As String
    
        strValue = txAcquisition(13)
        If F_InputBox.DoShow("Enter Value :", "Acquisition Value", , , , strValue, , True) Then
        
            txAcquisition(13) = strValue
        
        End If
        
        Exit Sub
    
    End If
    
    
    If rc Then


        If F_MsgBox.DoShow("Bottle already selected. Change number?", txAcquisition(0) & " | " & txStock(1)) Then
        
        Else
        
            Exit Sub
            
            
        End If
        
    Else
    
    
         
            
            
    End If
    
    Call ChangeBottleInAcquisition
    
End Sub




Private Function AddBottleInAcquisition(ByVal ID As Long) As Boolean
Dim rc As Boolean

On Error GoTo ERR_ADD

    rc = True

    If ID > 0 Then

        
        ReDim uBottle(0)
          
          Call GetDatabaseWareHouseEntry(ID, 0, True, uBottle)
          
          
          
          With uBottle(0)
          
             If .Finished <> "" Then
                MessageInfoTime = 2500
                PopupMessage 2, "This Bottle is Finished  : " & .Finished, , , .Lot & " | " & .EntryBottle
                
                
                 
                uBottle = MyWareHouseEntryCleanArray
                
                ReDim uBottle(0)
                            
                
                Call ClearAndFillAcquisition
                rc = False
                
                GoTo ERR_END
                
             End If
             
              txAcquisition(0) = .Lot
              txAcquisition(1) = .Location
              txAcquisition(2) = FormatDataLAT(.ArrivedTime)
              txAcquisition(3) = FormatDataLAT(.SupplierEXP)
              txAcquisition(4) = .Purity
              txAcquisition(5) = .Density
              txAcquisition(6) = FormatDataLAT(.Open)
              txAcquisition(7) = .MREXP
              txAcquisition(8) = GetStatus(.Status)
              txAcquisition(18) = .MRValueConcentration
              
              txStock(0) = .StockQTY
              txStock(1) = .EntryBottle
              
              lbUM = .Unit
              
          End With
    End If
    
    
    
ERR_END:
    On Error GoTo 0
    AddBottleInAcquisition = rc
    Exit Function
ERR_ADD:
    rc = False
    MsgBox Err.Description
    Resume Next
End Function


Private Function ChangeBottleInAcquisition() As Boolean
Dim ID As Long
Dim rc As Boolean
Dim strNumber As String
Dim strLot As String

On Error GoTo ERR_CHANGE

    rc = True
    
    strNumber = Trim(txStock(1))
    strLot = txAcquisition(0)
    
    If F_InputBox.DoShow("Enter Bottle Number :", "Bottle Number", , , "Select Bottle", strNumber) Then
    
        If strLot <> "" Then
        
            If F_MsgBox.DoShow("Change Lot number?", "Lot : " & strLot) Then
enterLot:
                If F_InputBox.DoShow("Enter Lot Number : ", "Lot Number", , , , strLot) Then
                
                
                End If
            End If
        Else
            GoTo enterLot
        End If
    Else
        frCommandInside_Click 19
        Exit Function
    End If
    
    If SearchIDBottle(uRecipe.Code, strNumber, strLot, ID) Then
    
    
        If ID = uBottle(0).ID Then
        Else
        
            AddBottleInAcquisition (ID)
        
        End If
    Else
        PopupMessage 2, "Didn't Find bottle!", , , strNumber & " | " & strLot
        
    End If
    
ERR_END:
    On Error GoTo 0
    ChangeBottleInAcquisition = rc
    Exit Function
ERR_CHANGE:
    rc = False
    MsgBox Err.Description
    Resume Next
   
End Function




Private Sub txAcquisition_Change(Index As Integer)
Dim rc As Boolean
    rc = IIf(Len(txAcquisition(Index)) > 0, True, False)
    
  '  txAcquisition(Index).BackColor = IIf(rc, vbWhite, &HF0F0F0)
    
    Select Case Index
        Case 0, 1
            'frCommandInside(9).Visible = rc
            'lbChemical = txAcquisition(0) & " | " & txAcquisition(1)
            
            'lbCritical = GetCriticalRM(txAcquisition(0))
            'lbNote = GetNoteRM(uRecipe.Code, txAcquisition(0))
            'lbNote.Visible = IIf(lbNote <> "", True, False)
            
            'If rc Then
            '    txAcquisition(3).BackColor = IIf(txAcquisition(3) <> "", vbWhite, &HF0F0F0)
            '    txAcquisition(4).BackColor = IIf(txAcquisition(4) <> "", vbWhite, &HF0F0F0)
            '    txAcquisition(5).BackColor = IIf(txAcquisition(5) <> "", vbWhite, &HF0F0F0)
            'End If
          
        Case 13
            ' peso acquisito
            
                 
            '---------------------------------------
            ' aggiorno lo STD e la bottiglia....
            '----------------------------------------
            UpdateFormAcquisition
        
            If Len(txAcquisition(0)) > 0 Then
                frCommandInside(8).Visible = rc
            Else
                 frCommandInside(8).Visible = False
            End If
            
            
            
            
    End Select
End Sub







Private Sub OpenComponentDatabase(ByVal rc As Boolean)
Dim userCode As String

   ' userCode = txFormulation(0)
   ' rc = FormChemicalRM.DoShow(userCode, , IIf(rc, HannaCode, ""))
   ' If rc Then
        
   '     If userCode <> "" Then
   '         Call FillUserBottle(userCode, False)
   '     End If
    
   ' End If
            
End Sub




Private Function CheckPreparationDetail() As Boolean
Dim rc As Boolean

    rc = True
    

    
    If txFormulation(1) = "" Then
        rc = False
        txFormulation(1).BackColor = vbRed
    End If
    
     If txFormulation(0) = "" Then
        rc = False
        txFormulation(0).BackColor = vbRed
    End If
       
    If txFormulation(8) = "" Then
        rc = False
        txFormulation(8).BackColor = vbRed
    End If
    
   
    CheckPreparationDetail = rc


End Function
Private Function AcquireAcquisition(ByVal rc As Boolean)
Dim mrc As Boolean
Dim i As Integer

    frCommandInside(5).Visible = False

    If CheckPreparationDetail = False Then
        frCommandInside(12).Visible = True
        ucScrollAdd1.UCScrollV.ScrollToValue 0
        PopupMessage 2, "Please fill all Preparation Details first..."
    
        Exit Function
    End If
    

    
    If rc Then
    
    
        If txAcquisition(16) > 0 Then
            If txAcquisition(0) = "" Or txStock(1) = "" Then
                PopupMessage 2, "Please select a bottle first..."
                frCommandInside_Click 19
                Exit Function
            End If
        End If

    If txAcquisition(13) = 0 And txAcquisition(10) > 0 Then
       PopupMessage 2, "Warning : Acquisition not Saved...", , True, lbInside(2)
       frCommandInside(8).Visible = False
       Exit Function
       
    ElseIf txAcquisition(13) = 0 And txAcquisition(10) = 0 And txAcquisition(16) <> 0 Then
    
        ' ho già acquisito questo STD....
        PopupMessage 2, "Please select a different Standard...", , , "Already Acquired"
        ucScrollAdd1.UCScrollV.ScrollToValue 0
        frInside(2).Visible = False
        Exit Function
        
        
    End If
        
        
        
    '---------------------------------------
    ' aggiorno lo STD e la bottiglia....
    '----------------------------------------
     Dim RealWeight As Double
    
     RealWeight = txAcquisition(13)
    
     Call SetSTDAfterAcquisition(RealWeight, uBottle(0))
        
        
    Call DoSaveAcquisition
        
        
    Else
       '
       
       
    End If
    
    ucScrollAdd1.UCScrollV.ScrollToValue 0
    frInside(2).Visible = False
    
End Function

Private Function DoSaveAcquisition()
Dim mrc As Boolean
    mrc = SaveAcquisition
    If mrc Then
        PopupMessage 2, "Acquisition done..."
    Else
        PopupMessage 2, "Warning : Acquisition not Saved...", , True, lbInside(2)
    End If
End Function


Private Function FillUserBottle(ByVal userCode As String, ByVal bFromBarcode As Boolean) As Boolean
Dim rc As Boolean
Dim Manufacurer As String
Dim ManufacturerCode As String

    On Error GoTo ERR_FILL:
    
     
    uBottle = MyWareHouseEntryCleanArray
    
    ReDim uBottle(0)
    
    Call ClearAndFillAcquisition
    
    frInside(2).Visible = True
    
    rc = True
    
ERR_END:
    On Error GoTo 0
    FillUserBottle = rc
    Exit Function
ERR_FILL:
    rc = False
    PopupMessage 2, Err.Description
    Resume ERR_END
End Function




Private Sub SetTxAcquisition()

Dim i As Integer

userAcquisition.bFromBarcode = False

Frame3(2).BackColor = &H886010
lbInside(2) = "STD : Acquisition Details"
lbInside(2).ForeColor = vbWhite
            
End Sub

Private Sub GetComponentSpecifics(ByVal userCode As String)



End Sub

Private Function ClearAcquisition()
Dim i As Integer
    
    For i = txAcquisition.LBound To txAcquisition.UBound
        txAcquisition(i) = ""
        txAcquisition(i).BackColor = &HF0F0F0
    Next
     For i = txStock.LBound To txStock.UBound
        txStock(i) = ""
    Next
    
    
  
    lbChemical = ""
    lbInside(2) = "Acquisition Details"
    PicTolerance.Visible = False
    
    frCommandInside(13).Visible = True
  
    Frame3(2).BackColor = &H886010
    lbInside(2) = "Acquisition Details"
    lbInside(2).ForeColor = vbWhite
    
    userAcquisition = userAcquisitionClean
    AcquisitionID = 0
    AcquisitionSTDNumber = ""
    AcquisitionWeight = 0
    frCommandInside(11).Visible = False
    
    SetTxAcquisition
    
End Function


Private Function SaveAcquisition(Optional bAcquisisciTutti As Boolean, Optional ByRef bBottigliaFinita As Boolean) As Boolean

Dim rc As Boolean



    If PreparationID = 0 Then
        ' devo prima salvare Preparaton in Database!!!!!
        Call SavePreparation
    End If
    

    rc = True
    
    '-------------------------------------------
    ' salvo l'acquisizione in userAcquisition
    ' registro i dati della bottiglia...
    '-------------------------------------------
    If uRecipe.STD(lSTDRow).Value <> 0 Then
    
    
        If txAcquisition(9) = "" Then
            
            PopupMessage 2, "Please Enter Pipette Code...", , , "STD preparation"
            txAcquisition_Click 9
            
           ' Exit Function
        
        End If

        userAcquisition.Bottle = uBottle(0).EntryBottle
        userAcquisition.MRLot = uBottle(0).Lot
    
        If uBottle(0).StockQTY < "0.1" Then
            
            ' chiudo la bottiglia...
            If F_MsgBox.DoShow("Close Bottle?", uBottle(0).Lot & " | " & uBottle(0).EntryBottle) Then
            
                uBottle(0).Finished = FormatDataLAT(Now())
        
                uBottle(0).Status = 2
                uBottle(0).PreparationID = PreparationID
                bBottigliaFinita = True
            End If
        Else
        
            If uBottle(0).Status = 0 Then
                Debug.Print uBottle(0).StockQTY
                uBottle(0).Open = FormatDataLAT(Now())
                uBottle(0).Status = 1
                uBottle(0).PreparationID = PreparationID
                bBottigliaFinita = False
            End If
                
            
        End If
        
        
        Call SaveWarehouseEntryInDatabase(uBottle(0), True, True)
        
    End If
    
    '-------------------------------------------
    ' salvo l'acquisizione in userAcquisition
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
    MaxCount = uRecipe.AcquisitionCount
    ReDim Preserve uRecipe.Acquisitions(MaxCount)
    
    uRecipe.Acquisitions(MaxCount) = userAcquisition
    
    uPreparation.Recipe = uRecipe
    
    Call RicaricoGrid1(True)
    
    SaveAcquisition = rc
    
   ' If uBottle(0).EntryBottle <> "" Then
   '     If bAcquisisciTutti = False Then
   '         If bBottigliaFinita = False Then
   '             If F_MsgBox.DoShow("Acquire All STDs left?") Then
   '
   '                 bAcquisisciTutti = True
   '
   '                 If CheckAllSTD Then
   '                 End If
   '
   '             End If
   '
   '         End If
   '     End If
   ' End If
    
      
End Function


Private Function CloseBottleInTable()



        ReDim uBottle(0)
          
          Call GetDatabaseWareHouseEntry(IDBottle, 0, True, uBottle)
        ' chiudo la bottiglia...
        If F_MsgBox.DoShow("Close Bottle?", uBottle(0).Lot & " | " & uBottle(0).EntryBottle) Then
        
            uBottle(0).Finished = FormatDataLAT(Now())
    
            uBottle(0).Status = 2
            uBottle(0).PreparationID = PreparationID
            
            Grid3.ReadOnly = False
            Grid3.Selection.DeleteByRow
            Grid3.ReadOnly = True
                
    

            Call SaveWarehouseEntryInDatabase(uBottle(0), True, True)
            
            IDBottle = 0
            
        End If
        

End Function
Private Function CheckAllSTD() As Boolean
Dim i As Integer
Dim rc As Boolean
Dim bBottigliaFinita As Boolean
With uRecipe

    rc = False
    If .STDcount > 0 Then
    
        For i = 1 To .STDcount
        
            lSTDRow = i
        
            If Abs(.STD(i).RealWeight - .STD(i).TheoreticalWeight) > TolerancePerc * .STD(i).TheoreticalWeight Then
                            
                rc = True
                MessageInfoTime = 2500
                
                
                PopupMessage 2, "Acquire for STD = " & .STD(i).NUMBER & vbCrLf & "STD Value = " & .STD(i).Value & vbCrLf & "Bottle Left = " & uBottle(0).StockQTY & uBottle(0).stockUnit
              
              
               
                
                Call fillAcquisitionDetail
                
                txAcquisition_Click 9
                 
                 
                Call AcquireMRfromBottle(uBottle(0))
                
                If uBottle(0).Status = 2 Then
                    ' close bottle!!!
                
                
                
                End If
                
                DoEvents
                
                bBottigliaFinita = False
                
                Call SaveAcquisition(True, bBottigliaFinita)
                
                If bBottigliaFinita Then
                    
                    ReDim uBottle(0)
                    ClearAndFillAcquisition
                
                    Exit Function
                
                End If
                
                DoEvents
            
            End If
        
        
        Next
    
    
    End If
    
    
End With

ucScrollAdd1.UCScrollV.ScrollToValue 0
frInside(2).Visible = False
    
CheckAllSTD = rc

End Function

Private Sub SetNewUserAcquisition()
Dim AcquisitionsCount As Integer

    With userAcquisition
    
        .AcquisitionTime = Now()
        .ActualWeight = txAcquisition(13)
        .LeftInBottle = uBottle(0).StockQTY & " " & uBottle(0).stockUnit
        .Bottle = uBottle(0).EntryBottle ' txStock(1)
        .Code = uRecipe.Code
        .DatePrep = uPreparation.DataPrep
        .FileName = SettingName
        .HannaCode = uPreparation.HannaCode.Code
        .HourPrep = uPreparation.HourPrep
       ' .MotherSolutionDate=
        .MRLot = uBottle(0).Lot ' txAcquisition(0)
        .MsType = uPreparation.MsType
        .Note = txAcquisition(17)
        .Operator = MyOperatore.Name
        .PreparationID = PreparationID
        .STDNumber = txAcquisition(15)
    
        .STDUnit = txAcquisition(19)
        .STDQty = txAcquisition(10)
        .STDValue = txAcquisition(16)
        .WeekPrep = uPreparation.PrepWeek
        .CodicePipetta = txAcquisition(9)
        
        .MotherSolutionDate = uPreparation.MotherSol.DataMS

    End With
    
    
    AcquisitionsCount = uRecipe.AcquisitionCount
    AcquisitionsCount = AcquisitionsCount + 1
    uRecipe.AcquisitionCount = AcquisitionsCount

    userAcquisition.Index = AcquisitionsCount

    uPreparation.Recipe = uRecipe
    
    
End Sub



Private Function SaveAcquisitionInTabAcquisition()

With dbTabAcquisition
    .AddNew
    !AcquisitionTime = userAcquisition.AcquisitionTime
    !Code = userAcquisition.Code
    !Bottle = userAcquisition.Bottle
    !MRLot = userAcquisition.MRLot
    !Index = userAcquisition.Index
    !ActualWeight = userAcquisition.ActualWeight
    !HannaCode = userAcquisition.HannaCode
    !LeftInBottle = userAcquisition.LeftInBottle
    !WeekPrep = uPreparation.PrepWeek
    !HourPrep = uPreparation.HourPrep
    !DatePrep = uPreparation.DataPrep

    !Note = userAcquisition.Note
    !Operator = userAcquisition.Operator
    !MsType = userAcquisition.MsType
    
   !STDNumber = userAcquisition.STDNumber
   !STDValue = userAcquisition.STDValue
   !STDQty = userAcquisition.STDQty
   !STDUnit = userAcquisition.STDUnit
   !MotherSolutionDate = IIf(IsNull(userAcquisition.MotherSolutionDate) Or userAcquisition.MotherSolutionDate = "", 0, userAcquisition.MotherSolutionDate)
    
    !CodicePipetta = userAcquisition.CodicePipetta
    
    !FileName = SettingName
    !PreparationID = PreparationID
    
   
    .Update
    
    userAcquisition.ID = !ID

End With


End Function



Private Function AddSTD() As Boolean
Dim strSTD As String
Dim i As Integer
Dim t As Integer
Dim Count As Long

With Grid1
    .AutoRedraw = False
    Count = .Rows - 1
    
        If F_MsgBox.DoShow("Enter new Standard?") Then

            For lSTDCol = 1 To 2
            
again:
                strSTD = ""
                If F_InputBox.DoShow("Enter Value:", "STD :" & .Cell(0, lSTDCol).Text, , , , strSTD, , True) Then

                   
                
                    Select Case lSTDCol
                        
                        Case 1
                        
                            For t = 1 To .Rows - 1
                                
                                If .Cell(t, lSTDCol).Text = strSTD Then
                                    PopupMessage 2, "Value already in Table. Please enter new or modify"
                                    Exit Function
                                End If
                                
                               
                            Next
                            
                            .AddItem "", False
                            ReDim Preserve uRecipe.STD(.Rows - 1)
                            uRecipe.STD(.Rows - 1).NUMBER = Int(strSTD)
                            
                            .Cell(.Rows - 1, lSTDCol).Text = strSTD
                            .AutoRedraw = True
                            .Refresh
                            
                            
                        Case 2
                             For t = 1 To .Rows - 1
                                
                                If .Cell(t, lSTDCol).Text = strSTD Then
                                    PopupMessage 2, "Value already in Table. Please enter new or modify"
                                    GoTo again
                                End If
                            Next
                               
                            uRecipe.STD(.Rows - 1).NUMBER = CDbl(strSTD)
                            .Cell(.Rows - 1, lSTDCol).Text = strSTD
                            .AutoRedraw = True
                            .Refresh
                            
                    End Select
                
                End If
              
            Next
        End If
    .AutoRedraw = True
    .Refresh
End With


Call RicaricoGrid1


End Function


Private Function DeleteSTD()
Dim IndexComp As Integer

Dim i As Integer

If STDNumber <> "" Then

     
    If F_MsgBox.DoShow("Delete Standard : " & STDNumber & vbCrLf & "Value : " & Trim(PadString(UserSTDValue)), SelectedHannaCode, True) Then
    
    Else
        Exit Function
    End If


cont:

    
    '-----------------------------------
    ' cancello dalla tabella
    '-----------------------------------
    
    With Grid1
    
        .ReadOnly = False
        .Selection.DeleteByRow
        .ReadOnly = True
        
      
    End With

   
    Call RicaricoGrid1
End If

End Function


Private Function SetSTDFromGrid1()
Dim i As Integer
Dim SommaMR As Double
With Grid1

   If txFormulation(12) <> "" Then uPreparation.MS.Volume = txFormulation(12)
     
     
    If .Rows > 1 Then
    
       
      
        uRecipe.STDcount = .Rows - 1
        ReDim uRecipe.STD(.Rows - 1)
        
        For i = 1 To .Rows - 1

            uRecipe.STD(i).NUMBER = .Cell(i, 1).Text
            uRecipe.STD(i).Value = IIf(.Cell(i, 2).Text = "", "0", .Cell(i, 2).Text)
            uRecipe.STD(i).Unit = .Cell(i, 3).Text
            uRecipe.STD(i).TheoreticalWeight = IIf(IsNull(.Cell(i, 4).Text) Or .Cell(i, 4).Text = "", 0, .Cell(i, 4).Text)
            uRecipe.STD(i).RealWeight = IIf(IsNull(.Cell(i, 5).Text) Or .Cell(i, 5).Text = "", 0, .Cell(i, 5).Text)
            uRecipe.STD(i).Variance = IIf(IsNull(.Cell(i, 6).Text) Or .Cell(i, 6).Text = "", 0, Replace(.Cell(i, 6).Text, "%", ""))
         
            
            SommaMR = SommaMR + uRecipe.STD(i).Value
        Next
        
        uRecipe.TotalWeight = SommaMR
        uRecipe.STDcount = .Rows - 1

    End If
End With

uPreparation.Recipe = uRecipe




End Function


Private Function RicaricoGrid1(Optional ByVal bNotSetSTD As Boolean)
    
Dim bAcquisitiTutti As Boolean

    frCommandInside(14).Visible = False
    
    uPreparation.HannaCode.STDVolume = txFormulation(4)

    '-------------------------------------------
    ' Copio i valori di STD dalla Grid1
    '-------------------------------------------
    
    If bNotSetSTD Then
    Else
        Call SetSTDFromGrid1
    End If
    '-------------------------------------------
    ' ricarico la Gird1
    '-------------------------------------------
    
     
    uPreparation.Recipe = uRecipe
    
    Call FillGridPreparationFromFile(Grid1, uPreparation, 1, PreparationID, bAcquisitiTutti)
    
    frCommandInside(14).Visible = bAcquisitiTutti
    
    uRecipe.TotalWeight = uPreparation.Recipe.TotalWeight
    uRecipe.ActualWeight = uPreparation.Recipe.ActualWeight
    
  STDRefreshTotalMSVolume
    
  

   
End Function

Private Function STDRefreshTotalMSVolume()

  
    txFormulation(11) = FormatNumber(uRecipe.TotalWeight, 3)
    txFormulation(13) = FormatNumber(uPreparation.MS.Qty, 3)
    lbMRStock = uPreparation.HannaCode.MR.STOCK_UNIT
    
    ' controllo nel caso di MS se supero la qtà totale che volevo preparare...
    
    Call CheckMS(uRecipe.TotalWeight, txFormulation(12))
    
    
    AggiornaMotherSolutionVolume

End Function


Private Function CheckMS(ByVal TotalWeight As Double, ByRef MSVolume As TextBox) As Boolean
' controllo nel caso di MS se supero la qtà totale che volevo preparare...
With uPreparation
If MSVolume <> "" Then
    If TotalWeight > MSVolume Then
        ' attenzione ho superato!!!!!
        
        
        If .MotherSol.Code = "" Then
            ' caso 1 non ho ancora MS
            ' modifico MRVolume e continuo....
            .MS.Volume = TotalWeight
            
            MSVolume = TotalWeight
            MSVolume.BackColor = vbColorAzzurrino ' vbColorOrange
          '  Image2_Click
        
        Else
            ' caso 2 ho già MS
        
        End If
    
    End If
End If
End With

End Function



Private Function DeleteAcquisition()
Dim IndexComp As Integer

Dim i As Integer

If AcquisitionID > 0 And AcquisitionSTDNumber <> "" Then

     
    If F_MsgBox.DoShow("Delete Acquisition STD : " & AcquisitionSTDNumber & vbCrLf & "Weight : " & Trim(PadString(AcquisitionWeight)) & "g", SelectedHannaCode, True) Then
    
    Else
        Exit Function
    End If
    '-----------------------------------
    ' sottraggo il peso inserito...
    '-----------------------------------
    With uRecipe
    
      '  Debug.Print .Acquisitions(lAcquisitionRow).ActualWeight
        .Acquisitions(lAcquisitionRow).bDeleted = True
        
        
        For i = 1 To .STDcount
        
            If .STD(i).Value = AcquisitionSTDNumber Then
                
                ' sottraggo...
                
                If .STD(i).RealWeight < AcquisitionWeight Then
           
                    .STD(i).RealWeight = 0
                    
                Else
                    .STD(i).RealWeight = .STD(i).RealWeight - AcquisitionWeight
                    
                End If
                GoTo cont:
            End If
        
        Next
    
    End With
    
    If F_MsgBox.DoShow("Standard not found in Recipe. Delete acquisition anyway?", AcquisitionSTDNumber) Then
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
    
    
    Call RicaricoGrid1(True)
    
End If

End Function


Private Function SavePreparation()
Dim i As Integer



    If CheckPreparationDetail = False Then
    
        PopupMessage 2, "Please fill Preparation Details first...."
        Exit Function
    
    End If
    
    If SettingName = "" Then SetSettingName



    lbWait = "Wait : Saving Data | Preparation..."
    lbHannaCode.Visible = False
    lbWait.Visible = True
    
    
    

    '---------------------------------------------
    ' aggiorna RealPerc / Variance / VariancePerc
    '---------------------------------------------
    Call SetSTDFromGrid1
    Call SetPreparationPercentageInSTD

    With uPreparation
        .Recipe = uRecipe
     
        SetHannaCodeDataPreparation
     
        
    End With
    
    '-------------------------------------------
    ' Salva e aggiorna TabPreparation
    '-------------------------------------------
    Call AggiornaTabPreparation(PreparationID, uPreparation)
   
    '-------------------------------------------
    ' Salva e aggiorna File
    '-------------------------------------------
    Call PeparationSaveSetting(uPreparation, SettingName)
    
    lbWait.Visible = False
    lbHannaCode.Visible = True
    
    PopupMessage 2, "Preparation correctly Saved", , , uRecipe.Code

    
End Function


Private Function SetPreparationPercentageInSTD()

Dim NewTotal As Double
Dim Perc     As Double
Dim Variance As Double
Dim VariancePerc As Double
Dim i As Integer

    With uRecipe
            If .ActualWeight <> 0 Then
            For i = 1 To .STDcount
              
              If .STD(i).RealWeight <> 0 Then
              
                    If (.STD(i).RealWeight) > 0 And (.STD(i).TheoreticalWeight) > 0 Then
                        .STD(i).Variance = .STD(i).TheoreticalWeight - .STD(i).RealWeight
                    End If
                    
                    If (.STD(i).Variance) <> 0 And (.STD(i).TheoreticalWeight) > 0 Then
                         .STD(i).VariancePerc = (.STD(i).Variance / .STD(i).TheoreticalWeight) * 100
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


txQRCode = ""
txQRCode.Top = 0
txQRCode.Visible = True
txQRCode.SetFocus

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
    

        
        If FillUserBottle(UserQrCode.Code, True) Then
        
        
            ' se è un componente o se ho comunque deciso di acquisirlo allora
            ' copio i dati in form
            ' altrimentri niente...
            
        
            txAcquisition(0) = UserQrCode.Code
            
            
           
            txQRCode.Visible = True
            
            PopupMessage 2, "Read Code : " & UserQrCode.Code
        
            ' inserisci automaticamente la pesata...
            'txAcquisition_Click 13
            
        
            If bOpenProductClassificationAfterScan Then
                
                frCommandInside_Click 9
            
            End If
        
        
        
        End If
    Else
        PopupMessage 3, "Please Check QRCode or Scan Again...", , True, "QRCode"
    End If
    

End If

ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_QR:
    MessageInfoTime = 2000
    PopupMessage 2, Err.Description & vbCrLf & "Please repeat reading...", , , "QR Code Reader"
    Resume ERR_END:


End Sub


Private Sub SetClosedPreparation(ByVal rc As Boolean)

    
    frCommandInside(4).Visible = Not (rc)
    frCommandInside(0).Visible = Not (rc)
    frCommandInside(1).Visible = Not (rc)
    frCommandInside(5).Visible = False
    frCommandInside(10).Visible = Not (rc)
    frExcel.Move frCommandInside(4).Left, frCommandInside(4).Top, frCommandInside(4).Width, frCommandInside(4).Height
    frExcel.Visible = rc
End Sub

Private Sub frExcel_Click()
            
    Dim ExcelFilename As String
    ' export LOT Excel
    If SettingName = "" Then
    Else

        ExcelFilename = "PREP_" & FormatNomeFile(Trim(uPreparation.HannaCode.Code) & "." & Trim(uPreparation.MRCode) & "." & Trim(uPreparation.DataPrep) & "." & Trim(uPreparation.HourPrep))
        
        
       ''' SettingName = FormatNomeFile(Trim(uRecipe.HannaCode.Code) & "." & Trim(uRecipe.Code) & "." & txFormulation(1) & "." & txFormulation(2) & "." & txFormulation(3)) & "." & USER_ESTENSIONE_PREPARATION

        
        PopupMessage 2, "Exporting data to Excel : please wait...." & vbCrLf & ExcelFilename
        If Len(ExcelFilename) > 60 Then ExcelFilename = Left$(ExcelFilename, 59)
        Call EsportaPreparationExcel(SettingName, ExcelFilename, uPreparation)
    End If


End Sub


Private Sub SetuPreparationExpDate()
'If uRecipe.Exp = "" Then uRecipe.Exp = GetRecipeExp(uRecipe.Code)
'uPreparation.ExpDate = SetExpDate(uPreparation.DataPrep, uRecipe.Exp)
'txFormulation(8) = uPreparation.ExpDate
End Sub
















Private Sub Grid4_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)


NotesID = 0

    With Grid4
    
        If FirstRow > 0 Then
        
            NotesID = .Cell(FirstRow, 5).Text
            txRevision(0) = .Cell(FirstRow, 1).Text
            txRevision(1) = .Cell(FirstRow, 2).Text
            txRevision(2) = .Cell(FirstRow, 4).Text
            txRevision(3) = .Cell(FirstRow, 3).Text
            
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

End Select

End Sub




Private Sub ClearRevisionForm()



Dim i As Integer
For i = 0 To txRevision.UBound
    txRevision(i) = ""
Next
txRevision(2) = MyOperatore.Name

End Sub


Private Sub txRevision_Click(Index As Integer)
Dim userCode As String
Dim Answer As String
Dim Selected As String
Dim bNumber As Boolean
Dim sString As String
Dim rc As Boolean

    Selected = lbRevision(Index) ' "Preparation"
    Answer = txRevision(Index)
    sString = "Please Enter " & lbRevision(Index)
  
    bNumber = False
    
    If txRevision(2) = "" Then txRevision(2) = MyOperatore.Name
    
    cmbRevType.Visible = False
    
    Select Case Index
        Case 0
            If Answer = "" Then Answer = FormatDataLAT(Now())
        Case 1
            ' type
            cmbRevType.ZOrder
            cmbRevType.Visible = True
            Exit Sub
        Case 2
        
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

Private Function DeleteNotes(ByVal ID As Long) As Boolean
Dim rc As Boolean
Dim i As Integer

rc = True

For i = 1 To txRevision.UBound
    If txRevision(i) = "" Then
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

For i = 1 To txRevision.UBound
    If txRevision(i) = "" Then
        rc = False
        PopupMessage 2, "Please enter all fields...", , True, "Notes History"
        AddNotes = rc
        Exit Function
    End If
Next





With dbTabPreparationNotes
    .filter = ""
    .filter = "filename='" & SettingName & "' and NoteDate='" & txRevision(0) & "'"
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
        
        !NoteDate = txRevision(0)
        !Type = txRevision(1)
        !Description = IIf(Len(txRevision(3)) > 255, Left(txRevision(3), 255), txRevision(3))
        !Operator = txRevision(2)
        !FileName = SettingName
        !PreparationID = PreparationID
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
            PopupMessage 2, "File correcly created on Desktop", , , SelectedHannaCode & "_Note_History.xls"
        End If
    End If
End Sub


Private Sub cmbRevType_Click()
txRevision(1) = cmbRevType
cmbRevType.Visible = False
End Sub

Private Sub AddcmbRevType()


    With cmbRevType
        .AddItem "Revision"
        .AddItem "Improvement"
        .AddItem "Issue"
        .ListIndex = 0
    End With

End Sub



Private Function AcquireMRfromBottle(ByRef uBottle As WareHouseEntry)
Dim RealWeight As Double
Dim rc As Boolean
Dim BottleLeftQty As Double
Dim QtyforSTD As Double
Dim BottleQty As Double

Dim Um As String

    
    If uRecipe.STD(lSTDRow).Value = 0 Then
        
        ' questo è il caso 0
    
        txAcquisition(13) = 0
        frCommandInside_Click 8
        Exit Function
        
    End If
    
    
      
    ' attenzione alle um
    QtyforSTD = txAcquisition(10)
    BottleQty = uBottle.StockQTY
      
      
      
    rc = QtyforSTD < BottleQty
     
    Select Case rc
        
        Case True
            ' caso1
            ' acquisisco il peso corretto che manca : QtyforSTD
            ' se la somma dei pesi degli STD è minore della qty bottiglia , posso chiedere di automatizzare il processo.
            ' ciclo for per acquisire tutti gli STD
            ' controllo se la bottglia è da chiudere...

            RealWeight = QtyforSTD
            BottleLeftQty = BottleQty - QtyforSTD
            
        Case False
        
        ' caso 2
        ' non riesco ad acquisire tutto e mi rimane ancora Qty da acquisire > cambio la bottiglia
        ' devo chiudere la bottiglia , acquisire e scegliere una tra bottiglia
          
            RealWeight = BottleQty
            BottleLeftQty = 0
            
            
         
            
            
    End Select
    
    txAcquisition(13) = RealWeight

  '  txStock(0) = BottleLeftQty
    
    UpdateFormAcquisition
    
    
   
    
    Call SetVarianceAcquisition(lSTDRow)


End Function


Private Function UpdateFormAcquisition()

If uBottle(0).EntryBottle <> "" Then
    txStock(0) = FormatNumber(uBottle(0).StockQTY - txAcquisition(13), 3)
    
    Select Case txStock(0)
    
        Case Is < 2
            txStock(0).ForeColor = vbColorRed
            txStock(1).ForeColor = vbColorRed
        Case Is < 10
        
           txStock(0).ForeColor = vbColorOrange
            txStock(1).ForeColor = vbColorOrange
        Case Else
        
           txStock(0).ForeColor = &H886010
            txStock(1).ForeColor = &H886010
    End Select
End If

End Function

Private Function SetSTDAfterAcquisition(ByVal RealWeight As Double, ByRef uBottle As WareHouseEntry)
Dim STDsTotalQty As Double
Dim dblAcquisition As Double
Dim dblExtraAcq As String


dblAcquisition = txAcquisition(13)
dblExtraAcq = 0

If uBottle.EntryBottle <> "" Then

    If F_MsgBox.DoShow("Extra acquisition? (e.g. prime the pipette)", "Pipette Acquisition") Then
        If F_InputBox.DoShow("Enter numeric value (mL) :", "Extra Acquisition", , , , dblExtraAcq, , True) Then
            dblAcquisition = dblAcquisition + dblExtraAcq
        End If
    End If

    uBottle.StockQTY = FormatNumber(uBottle.StockQTY - dblAcquisition, 2)
    
    
    With uRecipe.STD(lSTDRow)
        
        .RealWeight = RealWeight
        .ActualWeight = txAcquisition(10)
        .ActualWeight = .ActualWeight - .RealWeight

        Call GetAllSTDTotalWeight(uRecipe, STDsTotalQty)
 
        txAcquisition(11) = .ActualWeight
        txAcquisition(13) = .RealWeight
        .RealWeight = FormatNumber(.TheoreticalWeight - .ActualWeight, 3)
        
    End With
    
    Call SetVarianceAcquisition(lSTDRow)
    
End If
    
End Function




Private Function SetVarianceAcquisition(ByVal i As Integer)

Dim Variance        As Double
Dim VariancePerc    As Double
Dim RealWeight      As Double
Dim ActualWeight    As Double
Dim bRecalculate    As Boolean
Dim bCorrection     As Boolean

  
    If txAcquisition(10) <> "" And txAcquisition(13) <> "" Then
      
         With uRecipe.STD(i)
    
             Variance = txAcquisition(10) - txAcquisition(13)
             VariancePerc = (Variance / .TheoreticalWeight) * 100
        
             txAcquisition(11) = PadString(Variance)
             txAcquisition(12) = FormatNumber(VariancePerc) & " %  "
           
             
                 Dim MyColor As OLE_COLOR
                 
                 MyColor = ColorTolerance(Variance, .TheoreticalWeight * (TolerancePerc), bRecalculate, bCorrection)
                 
                     PicTolerance.BackColor = MyColor
                     PicTolerance.Visible = True
         End With
    
    End If
        

End Function


Private Function ClosePreparation()



If F_MsgBox.DoShow("Close Preparation?", lbHannaCode) Then


uPreparation.Recipe = uRecipe
uPreparation.bClosed = True
uPreparation.CloseDate = FormatDataLAT(Now())


Call SavePreparation

Dim rc As Boolean

If PreparationID > 0 Then

    
     
    With dbTabPreparation
        .filter = ""
        .filter = "ID='" & PreparationID & "'"
        If .EOF Then
        
            
        
        Else
            !bClosed = True
            !CloseDate = FormatDataLAT(Now())
            
            bPreparationClosed = True
            
        End If
    End With
    
    End If
    
    
    PopupMessage 2, "Preparation Closed." & vbCrLf & "Creating Excel export file...", lbHannaCode


    frExcel_Click
 
 
 
If FileExists(USER_TEMP_PATH & SettingName) Then
    
    FileCopy USER_TEMP_PATH & SettingName, USER_DATA_PATH & SettingName

    Kill USER_TEMP_PATH & SettingName

End If


End If



End Function




Private Function CheckCloseBottle(ByRef uBottle As WareHouseEntry)


Dim rc As Boolean
    On Error GoTo ERR_CHECK
    
    rc = True

With uBottle
    If .StockQTY * Um(.stockUnit) < 0.1 Then
        ' è meglio chiudere la bottiglia....
        If F_MsgBox.DoShow("Close bottle?" & vbCrLf & "Qty = " & .StockQTY & .stockUnit, .EntryBottle) Then
        
            rc = True
            
            uBottle.Finished = FormatDataLAT(Now())
            uBottle.NumberBottle = 1
            uBottle.Status = 2
            
            PopupMessage 2, "Bottle closed...", , , .EntryBottle
        
        Else
            rc = False
            GoTo ERR_END:
        
        End If
    End If
End With

ERR_END:
    On Error GoTo 0
    CheckCloseBottle = rc
    
    Exit Function
ERR_CHECK:
    rc = False
   
    Resume ERR_END:


End Function













'---------------------------------------------------------------
'
'
'                       MOTHER SOLUTION
'
'
'---------------------------------------------------------------

Private Sub SetFrameMotherSolution(ByVal Code As String)
Dim rc As Boolean
    rc = IIf(Code <> "", True, False)
    IconMS.Visible = IIf(rc, True, False)
    txFormulation(12).BackColor = IIf(IconMS.Visible, vbColorGreen, vbWhite)
    txFormulation(13).BackColor = IIf(IconMS.Visible, vbColorGreen, vbWhite)
    
    txFormulation(12).ForeColor = IIf(IconMS.Visible, vbWhite, &H644603)
    txFormulation(13).ForeColor = IIf(IconMS.Visible, vbWhite, &H644603)
    lbMRStock = uPreparation.HannaCode.MR.STOCK_UNIT
          
End Sub
     
      


Private Function CreateMotherSolution(ByRef uBottle As WareHouseEntry) As Boolean
Dim rc As Boolean
Dim mrc As Boolean
Dim strNote As String
Dim strTotVolume As Double


If txFormulation(11) = "" Then Exit Function

With uPreparation


    
    '-----------------------------------------------------------------------------
    ' ora verifichiamo che nella bottiglia ci sia la Qty giusta di MR
    '-----------------------------------------------------------------------------
    
   If .MS.Qty <= uBottle.StockQTY Then
   
        '
        If F_MsgBox.DoShow("Prepare MSVolume = " & txFormulation(12) & " mL" & vbCrLf & "MRQty = " & txFormulation(13) & lbMRStock & vbCrLf & "With Bottle =" & uBottle.EntryBottle & " ?", "Mother solution") Then
        Else
            Exit Function
        End If
        
        If F_InputBox.DoShow("Enter Note", "Mother Solution : Notes", , , , strNote) Then
        
        End If
        
        '----------------------------------------------
        ' creo la Mother solution
        '----------------------------------------------
        
        .MotherSol.bClosed = False
        .MotherSol.Code = .MRCode
        .MotherSol.DataPrep = txFormulation(1)
        .MotherSol.HourPrep = txFormulation(0)
        .MotherSol.WeekPrep = txFormulation(8)
        .MotherSol.ExpDays = IIf(.HannaCode.MSEXP = "", 10, (.HannaCode.MSEXP))
        .MotherSol.DataMS = FormatDataLAT(Now())
        .MotherSol.DataExp = DateAdd("d", .MotherSol.ExpDays, .MotherSol.DataMS)
        .MotherSol.HannaCode = .HannaCode.Code
        .MotherSol.MsType = .MsType
        .MotherSol.Note = strNote
        .MotherSol.Operator = .Operator
        .MotherSol.QtyLeft = txFormulation(12)
        .MotherSol.DataMS = FormatDataLAT(Now())
        .MotherSol.QtyProduced = txFormulation(12)
        .MotherSol.QtyUsed = 0
        .MotherSol.PreparationID = PreparationID
        .MotherSol.Unit = "mL"
        
        .MotherSol.Bottle = uBottle
    
        ' in caso di MS1 devo verificare la reale purity della bottle di MR
        ' se è cambiata allora devo rivedere i parametri degli STD ( ricarcio grid1 )
        If uBottle.Purity <> "" Then
            If .HannaCode.MR.MRPurity <> uBottle.Purity Then
                Call AddPurityMR(uBottle.MRCode, uBottle.Purity)
                .HannaCode.MR.MRPurity = uBottle.Purity
                Call RicaricoGrid1(True)
            
            
            End If
        End If
        
         '----------------------------------------------
        ' Salvo MS in Database
        '----------------------------------------------
        
         Call SaveMotherSolutionInDatabase(.MotherSol)
        
        '----------------------------------------------
        ' aggiorno le info sulla bottiglia e salvo
        '----------------------------------------------
        
        If uBottle.Status = 0 Then
            uBottle.Open = FormatDataLAT(Now())
            uBottle.Status = 1
        End If
        
        uBottle.StockQTY = uBottle.StockQTY - .MS.Qty
        uBottle.PreparationID = PreparationID
        
        Call SaveWarehouseEntryInDatabase(uBottle, True, True)
        
        ucScrollAdd1.UCScrollV.ScrollToValue frInside(0).Top - 680
        
     
        SetFrameMotherSolution (.MotherSol.Code)
        PopupMessage 2, "Mother Solution Created Correctly", "Mother solution"
        
   Else
        PopupMessage 2, "Please select another Bottle"
    
   End If
     
     
     
End With



End Function



Private Function SetMotherSolution()


   If txFormulation(11) = "" Then Exit Function

   frInside(2).Visible = False
    
    
   With uPreparation.MotherSol

        If .Code <> "" Then
            
            '--------------------------------------------------
            ' ho già una bottiglia di MS selezionata
            '--------------------------------------------------
            If .bClosed Then
                GoTo cont:
            Else
            
                PopupMessage 2, "Mother Solution Already Prepared/Selected...."
                
                If F_MsgBox.DoShow("Do You need to CHANGE Mother Solution?", "Mother Solution", , "Change", "Exit") Then
                    GoTo cont
                Else
                
                   
                End If
                
            End If
            
        Else

            If PreparationID = 0 Then
                '--------------------------------------------------
                ' devo prima salvare Preparaton in Database!!!!!
                '--------------------------------------------------
                Call SavePreparation
            End If
        
cont:
        
            '-----------------------------------------------------------------------------
            ' attenzione devo controllare la Qty che deve essere >= alla somma degli STD
            '-----------------------------------------------------------------------------
            Dim strTotVolume As Double
            strTotVolume = CDbl(txFormulation(11))
            
            If uPreparation.MS.Volume < strTotVolume Then
                uPreparation.MS.Volume = strTotVolume
                Call AggiornaMotherSolutionVolume
            End If
            
            '-----------------------------------------------------------------------------
            ' scelgo se Crearla o prendera da magazzino....
            '-----------------------------------------------------------------------------
            If F_MsgBox.DoShow("Do You need to prepare Mother Solution?", "Mother Solution", , "Prepare", "Select") Then
                Call AcquisitionBottleMS
            Else
            
                Call SelectMotherSolutionFromDatabase
            End If
        End If
    End With
           

End Function

Private Function AggiornaMotherSolutionVolume()
'Call SetSTDFromGrid1
Call SetSTDTheoreticalWeight(uPreparation.MsType, uPreparation)
txFormulation(12) = uPreparation.MS.Volume
txFormulation(13) = FormatNumber(uPreparation.MS.Qty, 3)
lbMRStock = uPreparation.HannaCode.MR.STOCK_UNIT

End Function



Private Function GetMSAcquisition() As Boolean
Dim rc As Boolean

rc = True

    '---------------------------------
    ' controllo se è Value=0
    '---------------------------------
    
    If uRecipe.STD(lSTDRow).Value = 0 Then
            ' acquisisco immediatamente..
            ' ma cosa???
            frCommandInside(8).Visible = True
            If F_MsgBox.DoShow("STD Number : " & uRecipe.STD(lSTDRow).NUMBER & vbCrLf & "STD Value : " & uRecipe.STD(lSTDRow).Value & vbCrLf & "Acquire?", lbInside(2)) Then
    
                    Call SaveMSAcquisition(True)
            
            End If
            
        End If
        
    
    '---------------------------------
    ' controllo se è ho fatto MS
    '---------------------------------
    
    
    With uPreparation.MotherSol
        If .Code <> "" Then
            If .bClosed = True Then
                ' lho usata tutta....
                MessageInfoTime = 2500
                PopupMessage 2, "Mother solution Closed. Prepare Another Mother Solution..."
                
                
                
                
                
                Exit Function
                
            End If
            
        Else
        
            PopupMessage 2, "Please Make Mother Solution first..."
            Exit Function
        End If
    


End With


    '---------------------------------
    ' inserisco valore Acquired
    '---------------------------------
    
    ' total MS
    ' txFormulation (12)
    Dim MSAquired As String
    Dim QTyAquired As Double
    Dim strData As String
    strData = "STD " & uRecipe.STD(lSTDRow).NUMBER & " | VALUE = " & uRecipe.STD(lSTDRow).Value
    
    MSAquired = uRecipe.STD(lSTDRow).ActualWeight
    
    
    If F_InputBox.DoShow("Please Enter Value :", strData, , , , MSAquired, , True) Then
    
        If MSAquired = "" Then
Err:
            PopupMessage 2, "Please  enter a valid value"
            Exit Function
        
        End If
        
        If F_MsgBox.DoShow("Acquire " & MSAquired & " mL of Mother Solution?", strData) Then
            
            '----------------------------
            ' calcolo la rimanenza ecc..
            '-----------------------------
            QTyAquired = CDbl(MSAquired)
            Call AcquireMSfromBottle(QTyAquired, uPreparation.MotherSol)
            Call SaveMSAcquisition(True, , QTyAquired)
        
        Else
            Exit Function
        
        End If
    
    
    Else
    
        Exit Function
    End If


'---------------------------------
' inserisco Pipetta
'---------------------------------

'---------------------------------
' Salvo...
'---------------------------------

'----------------------------------------------------------------------
' controllo se non ho più STD > chiedo se chiudere la bottiglia di MS
'----------------------------------------------------------------------

GetMSAcquisition = rc

End Function


Private Function AcquireMSfromBottle(ByVal RealWeight As Double, ByRef MS As MotherSolution)
Dim rc As Boolean
Dim BottleLeftQty As Double
Dim QtyforSTD As Double
Dim BottleQty As Double
Dim dblAcquisition As Double
Dim dblExtraAcq  As String
Dim Um As String

With uRecipe.STD(lSTDRow)

    ' attenzione alle um
    QtyforSTD = .ActualWeight
    BottleQty = MS.QtyLeft
    
    QtyforSTD = FormatNumber(QtyforSTD - RealWeight, 3)
    
    .ActualWeight = QtyforSTD
    .RealWeight = FormatNumber(.TheoreticalWeight - .ActualWeight, 3)
    
    ' sottraggo il prelievo da MS Bottle
  
    

    dblAcquisition = RealWeight
    dblExtraAcq = 0

    
    If F_MsgBox.DoShow("Extra acquisition? (e.g. prime the pipette") Then
        If F_InputBox.DoShow("Enter numeric value (mL) :", "Mother Solution Extra Acquisition", , , , dblExtraAcq, , True) Then
            dblAcquisition = dblAcquisition + dblExtraAcq
        End If
    End If

    
    BottleQty = BottleQty - dblAcquisition
    
    
    
    MS.QtyLeft = FormatNumber(BottleQty, 2)
    
    txFormulation(12) = MS.QtyLeft
    
    rc = QtyforSTD < BottleQty
    Select Case rc
        Case True
        Case False
    End Select
End With


End Function





Private Function SaveMSAcquisition(Optional bAcquisisciTutti As Boolean, Optional ByRef bBottigliaFinita As Boolean, Optional ByVal RealWeight As Double) As Boolean

Dim rc As Boolean


With uPreparation.MotherSol

    If PreparationID = 0 Then
        ' devo prima salvare Preparaton in Database!!!!!
        Call SavePreparation
    End If
    

    rc = True
    
    
    '-------------------------------------------
    ' salvo l'acquisizione in userAcquisition
    ' registro i dati della bottiglia...
    '-------------------------------------------
    
    If uRecipe.STD(lSTDRow).Value <> 0 Then
    
         '-------------
        ' pipette....
        '-------------
       Dim strPipette As String
        
        
        If F_InputBox.DoShow("Enter Code :", "STD preparation : Pipette", , , , strPipette) Then
        
           userAcquisition.CodicePipetta = strPipette
        
        End If


        userAcquisition.Bottle = .Bottle.EntryBottle
        userAcquisition.MRLot = .Bottle.Lot
        
        If .QtyLeft < "0.5" Then
            
            ' chiudo la bottiglia...
            If F_MsgBox.DoShow("Close Mother Solution Bottle?") Then
            
                .bClosed = True
                .QtyUsed = .QtyProduced - .QtyLeft
                bBottigliaFinita = True
                PopupMessage 3, "Mother solution closed..."
            End If
        Else
        
            'If .Bottle.Status = 0 Then
            '    Debug.Print .Bottle.StockQTY
            '    .Bottle.Open = FormatDataLAT(Now())
            '    .Bottle.Status = 1
            '    .Bottle.PreparationID = PreparationID
            '    bBottigliaFinita = False
            'End If
                
            
        End If
        
        
       ' Call SaveWarehouseEntryInDatabase(.Bottle, True, True, .Bottle.ID)
        Call SaveMotherSolutionInDatabase(uPreparation.MotherSol, .ID)
        
    End If
    
End With
    '-------------------------------------------
    ' salvo l'acquisizione in userAcquisition
    '-------------------------------------------
    
    Call SetMSUserAcquisition(RealWeight)
    
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
    MaxCount = uRecipe.AcquisitionCount
    ReDim Preserve uRecipe.Acquisitions(MaxCount)
    
    uRecipe.Acquisitions(MaxCount) = userAcquisition
    
    uPreparation.Recipe = uRecipe
    
    Call RicaricoGrid1(True)
    
    SaveMSAcquisition = rc

End Function

Private Sub SetMSUserAcquisition(ByVal RealWeight As Double)
Dim AcquisitionsCount As Integer
Dim strNote As String

    
    If F_InputBox.DoShow("Enter Note:", "STD " & uRecipe.STD(lSTDRow).NUMBER & " Acquisition Note", , , , strNote) Then
        
    End If



    With userAcquisition
    
        .AcquisitionTime = Now()
        .ActualWeight = RealWeight
        .LeftInBottle = uPreparation.MotherSol.QtyLeft
        .Bottle = uPreparation.MotherSol.Bottle.EntryBottle
        .MRLot = uPreparation.MotherSol.Bottle.Lot
        .Code = uRecipe.Code
        .DatePrep = uPreparation.DataPrep
        .FileName = SettingName
        .HannaCode = uPreparation.HannaCode.Code
        .HourPrep = uPreparation.HourPrep
       ' .MotherSolutionDate=
        
        .MsType = uPreparation.MsType
        .Note = strNote
        .Operator = MyOperatore.Name
        .PreparationID = PreparationID
        .STDNumber = uRecipe.STD(lSTDRow).NUMBER
    
        .STDUnit = uPreparation.HannaCode.MeasurementUnit
        .STDQty = uRecipe.STD(lSTDRow).RealWeight
        .STDValue = uRecipe.STD(lSTDRow).Value
        .WeekPrep = uPreparation.PrepWeek
        
        .MotherSolutionDate = uPreparation.MotherSol.DataMS

    End With
    
    
    AcquisitionsCount = uRecipe.AcquisitionCount
    AcquisitionsCount = AcquisitionsCount + 1
    uRecipe.AcquisitionCount = AcquisitionsCount

    userAcquisition.Index = AcquisitionsCount

    uPreparation.Recipe = uRecipe
    
    
End Sub
















'---------------------------------------------------------------
'
'
'              MOTHER SOLUTION SELECT FROM DATABASE
'
'
'---------------------------------------------------------------



Private Function SelectMotherSolutionFromDatabase() As Boolean
Dim rc As Boolean
Dim MSVolume As Double

rc = True

    
    lMSRow = 0
    frCommandInside(24).Visible = False
    frCommandInside(22).Visible = False
    frCommandInside(26).Visible = False
    
    MSVolume = CDbl(txFormulation(12))
    Call FillMotherSolutionTable(Grid5, uRecipe.Code, MSVolume)
    frMotherTable.ZOrder
    frMotherTable.Visible = IIf(Grid5.Rows > 1, False, True)
    
    lbMotherTable = "No MS ready for " & uRecipe.Code
    
    frInside(5).Visible = True
    
    
    ucScrollAdd1.UCScrollV.ScrollToValue frInside(5).Top - 680
    
    SelectMotherSolutionFromDatabase = rc
    
    
End Function




Private Sub Grid5_DblClick()

If lMSRow > 0 Then

    Call MotherSolSelected

End If


End Sub




Private Sub Grid5_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
    lMSRow = 0
    frCommandInside(24).Visible = False
    frCommandInside(22).Visible = False
    frCommandInside(26).Visible = False
    
    If FirstRow > 0 Then
        lMSRow = FirstRow
      
    
    
    End If


End Sub



Private Function MotherSolSelected()
Dim rc As Boolean
Dim ID As Long

With uPreparation
    ID = Grid5.Cell(lMSRow, 11).Text
    
    rc = GetMotherSolutionFromDatabase(.MotherSol, ID)
    
    If rc Then
       
        ucScrollAdd1.UCScrollV.ScrollToValue frInside(0).Top - 680
          
        SetFrameMotherSolution (.MotherSol.Code)
        
        txFormulation(12) = .MotherSol.QtyLeft
        
        Image2_Click
        
        
        PopupMessage 2, "Mother Solution Selected", "Mother solution"
    End If
            
End With


End Function


Private Function GotoPipetteGrid()
            ' scroll to pipeette
            
            Dim Qty As Double
            
            
    With uPreparation
        Select Case .MsType
            
            Case 0
                Qty = IIf(IsNumeric(txAcquisition(10)), txAcquisition(10), 0)
                
            Case 1
                Qty = uRecipe.STD(lSTDRow).ActualWeight
            Case 2
                Qty = uRecipe.STD(lSTDRow).ActualWeight
        End Select
    End With
    
     frInside(6).Top = frInside(4).Top
      frInside(6).Left = frInside(4).Left
             frInside(4).Visible = False
            frInside(6).Visible = True
             frInside(6).ZOrder
            Call CaricaPipette(Grid6, Qty)

            
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(6).Top - 680
                  
End Function



          
            
