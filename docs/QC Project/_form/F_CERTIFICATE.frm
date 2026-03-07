VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form F_CERTIFICATE 
   BackColor       =   &H00886010&
   Caption         =   "Certificate Form"
   ClientHeight    =   12000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
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
   Icon            =   "F_CERTIFICATE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12510
   ScaleMode       =   0  'User
   ScaleWidth      =   19200
   Begin MSComDlg.CommonDialog dlgOption 
      Left            =   120
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicMenuBar 
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
      Index           =   0
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   19215
      TabIndex        =   5
      Top             =   0
      Width           =   19215
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   3840
         MouseIcon       =   "F_CERTIFICATE.frx":33E2
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   36
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   720
            MousePointer    =   99  'Custom
            Picture         =   "F_CERTIFICATE.frx":36EC
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Graph"
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
            Left            =   60
            MouseIcon       =   "F_CERTIFICATE.frx":559E
            MousePointer    =   99  'Custom
            TabIndex        =   37
            Top             =   720
            Width           =   1875
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   0
         MouseIcon       =   "F_CERTIFICATE.frx":58A8
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
            MouseIcon       =   "F_CERTIFICATE.frx":5BB2
            MousePointer    =   99  'Custom
            Picture         =   "F_CERTIFICATE.frx":5EBC
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Lot Certificate"
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
            Left            =   60
            MouseIcon       =   "F_CERTIFICATE.frx":88AE
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   720
            Width           =   1875
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   1920
         MouseIcon       =   "F_CERTIFICATE.frx":8BB8
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   6
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Lot Calculation"
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
            Left            =   60
            MouseIcon       =   "F_CERTIFICATE.frx":8EC2
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   720
            Width           =   1875
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   735
            MousePointer    =   99  'Custom
            Picture         =   "F_CERTIFICATE.frx":91CC
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Certificate of Analysis"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   12120
         TabIndex        =   10
         Top             =   240
         Width           =   6540
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   9855
      Index           =   0
      Left            =   0
      ScaleHeight     =   9855
      ScaleWidth      =   19215
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Timer TimerIntro 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   4320
         Top             =   240
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
         Left            =   3600
         TabIndex        =   42
         Top             =   9240
         Visible         =   0   'False
         Width           =   3015
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Create Certificate"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   43
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
         Index           =   0
         Left            =   480
         TabIndex        =   40
         Top             =   9240
         Width           =   3015
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Save Certificate"
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
            TabIndex        =   41
            Top             =   120
            Width           =   3015
         End
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
         Left            =   6720
         TabIndex        =   38
         Top             =   9240
         Visible         =   0   'False
         Width           =   3015
         Begin VB.Image Image 
            Height          =   480
            Left            =   120
            MousePointer    =   99  'Custom
            Picture         =   "F_CERTIFICATE.frx":C5AE
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
            TabIndex        =   39
            Top             =   120
            Width           =   3015
         End
      End
      Begin FlexCell.Grid Grd2 
         Height          =   6255
         Left            =   480
         TabIndex        =   21
         Top             =   600
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   11033
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColorActiveCellSel=   15790320
         BackColorBkg    =   16777215
         BackColorFixed  =   15790320
         BackColorFixedSel=   15790320
         BackColorScrollBar=   15592423
         BorderColor     =   16777215
         CellBorderColor =   15790320
         CellBorderColorFixed=   15790320
         Cols            =   5
         DefaultFontName =   "Century Gothic"
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         DefaultRowHeight=   20
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   16777215
         Rows            =   1
         ScrollBarStyle  =   0
         MultiSelect     =   0   'False
      End
      Begin FlexCell.Grid Grid1 
         Height          =   3495
         Left            =   10200
         TabIndex        =   23
         Top             =   600
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   6165
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColorActiveCellSel=   15790320
         BackColorBkg    =   16777215
         BackColorFixed  =   15790320
         BackColorFixedSel=   15790320
         BackColorScrollBar=   15592423
         BorderColor     =   16777215
         CellBorderColor =   15790320
         CellBorderColorFixed=   16777215
         Cols            =   5
         DefaultFontName =   "Century Gothic"
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         DefaultRowHeight=   20
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   16777215
         Rows            =   1
         ScrollBarStyle  =   0
         MultiSelect     =   0   'False
      End
      Begin FlexCell.Grid Grid2 
         Height          =   4575
         Left            =   10200
         TabIndex        =   25
         Top             =   4560
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   8070
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColorActiveCellSel=   15790320
         BackColorBkg    =   16777215
         BackColorFixed  =   15790320
         BackColorFixedSel=   15790320
         BackColorScrollBar=   15592423
         BorderColor     =   16777215
         CellBorderColor =   15790320
         CellBorderColorFixed=   16777215
         Cols            =   5
         DefaultFontName =   "Century Gothic"
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         DefaultRowHeight=   20
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   16777215
         Rows            =   1
         ScrollBarStyle  =   0
         MultiSelect     =   0   'False
      End
      Begin FlexCell.Grid Grid6 
         Height          =   1335
         Left            =   480
         TabIndex        =   46
         Top             =   7800
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   2355
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColorActiveCellSel=   15790320
         BackColorBkg    =   16777215
         BackColorFixed  =   15790320
         BackColorFixedSel=   15790320
         BackColorScrollBar=   15592423
         BorderColor     =   16777215
         CellBorderColor =   15790320
         CellBorderColorFixed=   16777215
         Cols            =   5
         DefaultFontName =   "Century Gothic"
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         DefaultRowHeight=   20
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   16777215
         Rows            =   1
         ScrollBarStyle  =   0
         MultiSelect     =   0   'False
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   2160
         Picture         =   "F_CERTIFICATE.frx":F990
         Top             =   200
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Components identification"
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
         Index           =   5
         Left            =   480
         TabIndex        =   47
         Top             =   6960
         Width           =   3045
      End
      Begin VB.Label lbReagent 
         Alignment       =   2  'Center
         BackColor       =   &H000080DF&
         Caption         =   "REAGENT SET 1"
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
         Index           =   50
         Left            =   480
         MouseIcon       =   "F_CERTIFICATE.frx":11402
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   7365
         Width           =   4215
      End
      Begin VB.Label lbReagent 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "REAGENT SET 2"
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
         Index           =   51
         Left            =   4800
         MouseIcon       =   "F_CERTIFICATE.frx":1170C
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   7365
         Width           =   4215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calibration Function"
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
         Left            =   10200
         TabIndex        =   26
         Top             =   4200
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Result"
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
         Left            =   10200
         TabIndex        =   24
         Top             =   240
         Width           =   5820
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
         Left            =   480
         TabIndex        =   22
         Top             =   240
         Width           =   1500
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Left            =   9360
         Picture         =   "F_CERTIFICATE.frx":11A16
         Top             =   8880
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   9855
      Index           =   2
      Left            =   0
      ScaleHeight     =   9855
      ScaleWidth      =   19215
      TabIndex        =   31
      Top             =   1080
      Width           =   19215
      Begin VB.PictureBox picGraph 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   14040
         ScaleHeight     =   795
         ScaleWidth      =   1275
         TabIndex        =   32
         Top             =   6960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSChart20Lib.MSChart chChart 
         Height          =   9135
         Left            =   2640
         OleObjectBlob   =   "F_CERTIFICATE.frx":14DF8
         TabIndex        =   33
         Top             =   360
         Width           =   13575
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   9855
      Index           =   1
      Left            =   0
      ScaleHeight     =   9855
      ScaleWidth      =   19215
      TabIndex        =   11
      Top             =   1080
      Width           =   19215
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         FillColor       =   &H00004000&
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   6600
         ScaleHeight     =   855
         ScaleWidth      =   5295
         TabIndex        =   13
         Top             =   8640
         Visible         =   0   'False
         Width           =   5295
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "CLOSE LOT"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   14
            Top             =   240
            Width           =   5295
         End
         Begin VB.Image ImageTAV 
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "F_CERTIFICATE.frx":1730E
            MousePointer    =   99  'Custom
            Picture         =   "F_CERTIFICATE.frx":17618
            Top             =   180
            Width           =   480
         End
      End
      Begin FlexCell.Grid Grid3 
         Height          =   9015
         Left            =   240
         TabIndex        =   27
         Top             =   600
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   15901
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColorActiveCellSel=   15790320
         BackColorBkg    =   16777215
         BackColorFixed  =   15790320
         BackColorFixedSel=   15790320
         BackColorScrollBar=   15592423
         BorderColor     =   16777215
         CellBorderColor =   15790320
         CellBorderColorFixed=   16777215
         Cols            =   5
         DefaultFontName =   "Century Gothic"
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         DefaultRowHeight=   20
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   16777215
         Rows            =   1
         ScrollBarStyle  =   0
         MultiSelect     =   0   'False
      End
      Begin FlexCell.Grid Grid4 
         Height          =   2175
         Left            =   6600
         TabIndex        =   29
         Top             =   600
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   3836
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColorActiveCellSel=   15790320
         BackColorBkg    =   16777215
         BackColorFixed  =   15790320
         BackColorFixedSel=   15790320
         BackColorScrollBar=   15592423
         BorderColor     =   16777215
         CellBorderColor =   15790320
         CellBorderColorFixed=   16777215
         Cols            =   5
         DefaultFontName =   "Century Gothic"
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         DefaultRowHeight=   20
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   16777215
         Rows            =   1
         ScrollBarStyle  =   0
         MultiSelect     =   0   'False
      End
      Begin FlexCell.Grid Grid5 
         Height          =   5175
         Left            =   6600
         TabIndex        =   34
         Top             =   3360
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   9128
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColorActiveCellSel=   15790320
         BackColorBkg    =   16777215
         BackColorFixed  =   15790320
         BackColorFixedSel=   15790320
         BackColorScrollBar=   15592423
         BorderColor     =   16777215
         CellBorderColor =   15790320
         CellBorderColorFixed=   16777215
         Cols            =   5
         DefaultFontName =   "Century Gothic"
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         DefaultRowHeight=   20
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   16777215
         Rows            =   1
         ScrollBarStyle  =   0
         MultiSelect     =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calibration uncertainty"
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
         Left            =   6600
         TabIndex        =   35
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Linear Regression  (y = a + bx)"
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
         Index           =   4
         Left            =   6600
         TabIndex        =   30
         Top             =   240
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Curve Data"
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
         Index           =   3
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Specifications Table"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   7800
         MouseIcon       =   "F_CERTIFICATE.frx":1A9FA
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   960
         Width           =   3255
      End
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   3
      Left            =   3960
      MouseIcon       =   "F_CERTIFICATE.frx":1AD04
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   10800
      Width           =   1935
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   1
      Left            =   8280
      MouseIcon       =   "F_CERTIFICATE.frx":1B00E
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   10680
      Width           =   2775
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   0
      Left            =   17640
      MouseIcon       =   "F_CERTIFICATE.frx":1B318
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   10800
      Width           =   1695
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   2
      Left            =   15240
      MouseIcon       =   "F_CERTIFICATE.frx":1B622
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   10800
      Width           =   2295
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   4
      Left            =   0
      MouseIcon       =   "F_CERTIFICATE.frx":1B92C
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   10920
      Width           =   1935
   End
   Begin VB.Label La 
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
      Index           =   2
      Left            =   17655
      MouseIcon       =   "F_CERTIFICATE.frx":1BC36
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   11715
      Width           =   1200
   End
   Begin VB.Label La 
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
      Index           =   1
      Left            =   15630
      MouseIcon       =   "F_CERTIFICATE.frx":1BF40
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   11715
      Width           =   1230
   End
   Begin VB.Label La 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Procedure"
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
      Left            =   9000
      MouseIcon       =   "F_CERTIFICATE.frx":1C24A
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   11715
      Width           =   1200
   End
   Begin VB.Label La 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Excel Folder"
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
      Left            =   4290
      MouseIcon       =   "F_CERTIFICATE.frx":1C554
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   11715
      Width           =   990
   End
   Begin VB.Label lbOperator 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
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
      Left            =   300
      TabIndex        =   16
      Top             =   11715
      Width           =   645
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   4
      Left            =   360
      MouseIcon       =   "F_CERTIFICATE.frx":1C85E
      MousePointer    =   99  'Custom
      Picture         =   "F_CERTIFICATE.frx":1CB68
      Top             =   11160
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   15960
      MouseIcon       =   "F_CERTIFICATE.frx":1FF4A
      MousePointer    =   99  'Custom
      Picture         =   "F_CERTIFICATE.frx":20254
      Top             =   11160
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_CERTIFICATE.frx":23636
      Height          =   480
      Index           =   1
      Left            =   9360
      MouseIcon       =   "F_CERTIFICATE.frx":26A18
      MousePointer    =   99  'Custom
      Picture         =   "F_CERTIFICATE.frx":26D22
      Top             =   11160
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   4560
      MouseIcon       =   "F_CERTIFICATE.frx":2A104
      MousePointer    =   99  'Custom
      Picture         =   "F_CERTIFICATE.frx":2A40E
      Top             =   11160
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   431.698
      X2              =   16836.23
      Y1              =   11133.9
      Y2              =   11133.9
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   12950.94
      X2              =   12950.94
      Y1              =   125.1
      Y2              =   12510
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   4316.981
      X2              =   4316.981
      Y1              =   125.1
      Y2              =   12510
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   8633.963
      X2              =   8633.963
      Y1              =   250.2
      Y2              =   12635.1
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_CERTIFICATE.frx":2D7F0
      Height          =   480
      Index           =   0
      Left            =   18240
      MouseIcon       =   "F_CERTIFICATE.frx":30BD2
      MousePointer    =   99  'Custom
      Picture         =   "F_CERTIFICATE.frx":30EDC
      Top             =   11160
      Width           =   480
   End
End
Attribute VB_Name = "F_CERTIFICATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private IndexFormProcedura As Integer
Private IndexMainProcedura As Integer
Private IndexTextSelected As Integer
Private IndexText As Integer
Private MyLot As String

Private MyCodeID As String
Private MyCode As String
Private m_rc As Boolean
Private bFormSaved As Boolean
Private MeasurementUnit As String
Private UserDecimal As String
Private intDecimal As Integer
Private STD() As String
Private MeterNumber As Integer
Private pHNumber As Integer
Private pHMin(3) As String
Private pHMax(3) As String
Private MinNumber80Pec(10) As Integer
Private SelectedSTDNumber As Integer
Private strPassed As String
Private numSelectedStandard As String
Private AndOr As String

Private lCol As Long
Private lRow As Long

Private bAnotherFormCalled As Boolean

Private IndexProcedura As Integer
Private IndexDashCommInside As Integer
Private IndexVisibleFrame As Integer

'------------------------------------------------------------
'               GRAFICO
'------------------------------------------------------------


'these 6 Private statement are needed for saving graph
Private Declare Function SendMessage Lib "user32" Alias _
  "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_PAINT = &HF
Private Const WM_PRINT = &H317
Private Const PRF_CLIENT = &H4&    ' Draw the window's client area.
Private Const PRF_CHILDREN = &H10& ' Draw all visible child windows.
Private Const PRF_OWNED = &H20&    ' Draw all owned windows.

Private oRed As ColorConstants
Private oGreen As ColorConstants
Private oBlue As ColorConstants


Private MyChemicalQC As ptChemicalQC
Private myCertificate As CertType

Private bAllValue As Boolean
Private STDCount As Integer


Private Virgola As Integer

Private UserRow As Long


Private ExcelApp As Object
Private NumLinearità As Integer


Private iReagentSet() As RegSet

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
            .Left = Ctl.X1
            .Top = Ctl.Y1
            .Width = Ctl.X2 - Ctl.X1
            .Height = Ctl.Y2 - Ctl.Y1
        ElseIf TypeOf Ctl Is Menu Then
        ElseIf TypeOf Ctl Is Inet Then
        ElseIf TypeOf Ctl Is Timer Then
        ElseIf TypeOf Ctl Is CommonDialog Then
        

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
            Ctl.X1 = x_scale * .Left
            Ctl.Y1 = y_scale * .Top
            Ctl.X2 = Ctl.X1 + x_scale * .Width
            Ctl.Y2 = Ctl.Y1 + y_scale * .Height
        ElseIf TypeOf Ctl Is Timer Then
        ElseIf TypeOf Ctl Is Inet Then
        ElseIf TypeOf Ctl Is CommonDialog Then
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
Public Function DoShow(ByRef Index As Integer, Optional ByRef sLot As String, Optional ByRef sCode As String, Optional ByVal lngID As Long, Optional MyImage As Image, Optional FileName As String) As Boolean

    On Error GoTo ERR_SHOW
    
    Dim rc As Boolean
    'Set DefaultMenu(4) = MyImage
    IndexMainProcedura = Index
    m_rc = False
    bFormSaved = False
    
    MyCode = sCode
    SettingName = FileName
    MyCodeID = lngID
    
    
    
     
    MyChemicalQC = MyChemicalQCClean
    myCertificate = MyCertificateClean

    
    Call SetGridFGCodeLotCalculation(Grid3)
    Call SetGridFGCodeCertificate(Grd2)
    Call SetGridFGCodeCalFunction(Grid2)
    Call SetGridFGCodeLotResult(Grid1)
    
    Call SetGridFGCodeLotCalculationLinearRegression(Grid4)
    Call SetGridCalibrationUncertainty(Grid5)
    Call SetGridReagentSet(Grid6)


    rc = GetSTDValueForGraph
    
    
    GetFormSettingName
    
    If Not (rc) Then
        Exit Function
    End If
    
    
    
       
    If sLot <> "" And sCode <> "" Then
        'Call GetCodeInformation(sLot, sCode, lngID)
    Else
        PopupMessage 2, "Please select a valid FGCode..."
        Unload Me
    End If
    

    blTable = sCode & " | Certificate of Analysis"

    
    
    Call CopyFGCodeCertificateGrd2(Grd2, lngID, myCertificate)
    Call CopyFGCodeCertificateGrid2(Grid2, lngID, myCertificate.CalibrationFunction)
   
    
    SelectProcedura 0
    mOk
    
    
    If MyOperatore.IndexPrivilege >= 1 Then
            
        'Text1(11).Locked = False
        'Text1(12).Locked = False
        'Text1(12) = MyOperatore.Name
    
    End If
    
    TimerIntro.Enabled = True
  

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
    MsgBox Err.Description
    Resume ERR_END
End Function



Private Function LotCalculations()

 
    Call GetCertificateDescriptionFromFile
    
    Call CopyLotResultSTD

  
    Call SetLotCalculation
    
    
    lbReagent_Click 50
    

End Function


Private Sub DefaultMenuLabel_Click(Index As Integer)
Dim MyIndex As Integer
Select Case Index
    Case 0
        ' vai avanti
        If IndexFormProcedura = 1 Then
            MyIndex = 0
        Else
            MyIndex = IndexFormProcedura + 1
        End If
        PicMenu_Click MyIndex
    Case 1
    
        Unload Me
      
    Case 2
        ' torna indietro
        If IndexFormProcedura = 0 Then
            MyIndex = 1
        Else
            MyIndex = IndexFormProcedura - 1
        End If
        PicMenu_Click MyIndex
    Case 3
         ApriIlReportFolder (USER_EXCEL_PATH)
    Case 4
         frmLogin.DoShow 1
      
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
lbOperator = MyOperatore.Name

SaveSizes
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37
        DefaultMenuLabel_Click 2
    Case 39
        DefaultMenuLabel_Click 0
End Select
End Sub



Private Sub Form_Resize()

ResizeControls
End Sub




Private Sub frCommandInside_Click(Index As Integer)
Dim rc As Boolean
Dim bVaue As Boolean
Dim MyExcelName As String
Dim FileSavedName As String
Select Case Index
    Case 0
        Call SaveMe
        Call SetLotCalculation
    Case 1
        bVaue = IIf(Grid6.cell(2, 1).text <> "", True, False)
        
        FileSavedName = FormatNomeFile(Grd2.cell(2, 2).text & "." & Grd2.cell(5, 2).text)
        
       ' If bVaue Then
            rc = SetExcelCertificate_NEW(myCertificate, SettingName, MyExcelName, FileSavedName)
       ' Else
          '  rc = SetExcelCertificate(MyCertificate, SettingName, MyExcelName)
        
       ' End If
        If rc Then
            PopupMessage 2, "Excel and pdf files correctly generated.", , , "Lot: " & myCertificate.LotNumber
        End If
    Case 2
      
End Select


End Sub

Private Sub frCommandInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    
IndexDashCommInside = Index
Dim i As Integer
    For i = 0 To frCommandInside.UBound
        If i = Index Then
            ' quando ci passo sopra....
            frCommandInside(i).BackColor = &H846623
            lbCommandInside(i).ForeColor = vbWhite
            If i = 0 Then
                frCommandInside(i).BackColor = &H20A020
            End If
            
        Else
            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 0 Then
                frCommandInside(i).BackColor = &H8000&
            End If
        End If
    
    Next
 
 
End Sub

Private Sub frExcel_Click()

'Grid2.ExportToExcel USER_DESKTOP & "\" & FormatNomeFile(FGCode) & "_RevHistory.xls", True, True
'DoEvents
'MessageInfoTime = 2500
'PopupMessage 2, "File correcly created on Desktop", , , FormatNomeFile(FGCode) & "_RevHistory.xls"

End Sub



Private Sub Grd2_CellChange(ByVal Row As Long, ByVal Col As Long)
    Select Case Row
        Case 15
            If Col = 2 Then
                With myCertificate
                    .RangeFormula = Grd2.cell(15, 2).text
                    
                    Label1(0) = "Lot Result [" & .RangeFormula & "]"
                    
                    If Grid2.Rows > 1 Then
                        
                        Grid2.cell(2, 0).text = "Standard Deviation [" & .RangeFormula & "]"
                        Grid2.cell(4, 0).text = "Confidence interval (95%)[" & .RangeFormula & "]"
                        Grid2.AutoFitRowHeight (2)
                        Grid2.AutoFitRowHeight (4)
                    End If
                End With
        
            End If
    
    End Select
End Sub

Private Sub Grd2_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

UserRow = 0
     Grd2.ReadOnly = True
    
    If FirstRow > 0 Then
        UserRow = FirstRow
        
    
    End If
    
    
End Sub

Private Sub Grd2_OwnerDrawCell(ByVal Row As Long, ByVal Col As Long, ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Handled As Boolean)
 Dim rc As Boolean
    
    Select Case Row
   
   '
        Case 5, 6, 7
            If Col = 2 Then   ' visualizzo il pulsante Salva!
                rc = IIf(Grd2.cell(5, 2).text <> "", True, False)
                rc = IIf(Grd2.cell(6, 2).text <> "", rc, False)
                rc = IIf(Grd2.cell(7, 2).text <> "", rc, False)
                
                
                frCommandInside(0).Visible = rc


                
    
            End If
    End Select
    
End Sub
Private Sub Grd2_Click()
Dim str As String
Dim bNextControllo As Boolean

Dim RepStr As String

        
        
         '.Cell(1, 1).Text = "Product Name"
         '.Cell(2, 1).Text = "Product Code"
         '.Cell(3, 1).Text = "Method"
         '.Cell(4, 1).Text = "Range ppm (as O2)"
         '.Cell(5, 1).Text = "Lot number"
         '.Cell(6, 1).Text = "Best use before"
         '.Cell(7, 1).Text = "Date of analysis"
         '.Cell(8, 1).Text = "Reference meter"
         '.Cell(9, 1).Text = "Reference standard"
         '.Cell(10, 1).Text = "Wavelenght nm"
         '.Cell(11, 1).Text = "Cell mm"
         '.Cell(12, 1).Text = "Reference standard Note"
         
         
      
   
    str = Grd2.cell(UserRow, 1).text
    RepStr = Grd2.cell(UserRow, 2).text
    Select Case UserRow
    
        Case 1
            
        
        Case 10
            ' valori numerici
            If F_InputBox.DoShow("Enter Requested Value and press Save", str, , , , RepStr, , True) Then
                Grd2.cell(UserRow, 2).text = RepStr
            End If

        Case 6, 7
            ' date
            If RepStr = "" Then RepStr = FormatDataLAT(Now())
            If F_InputBox.DoShow("Enter a Valid Date", str, , , , RepStr) Then
                If IsDate(RepStr) Then
                
                    Grd2.cell(UserRow, 2).text = FormatDataLAT(RepStr)
                
                End If
            End If
     
            
            
        Case Else
            ' strighe generiche
             If F_InputBox.DoShow("Enter Requested Value and press Save", str, , , , RepStr) Then
                Grd2.cell(UserRow, 2).text = RepStr
            End If
    
    
    End Select

    If UserRow = 18 Then Grd2.AutoFitRowHeight (UserRow)
    Grd2.Refresh
    
  '  SetLeakingLotVariables

End Sub




Private Sub Grid6_Click()
Dim sString As String
Dim sResult As String
    With Grid6
        If lRow > 0 Then
            If lCol > 0 Then
                sString = .cell(lRow, 0).text
                sResult = .cell(lRow, lCol).text
                If F_InputBox.DoShow("Enter Value", sString & " #" & lCol, , , , sResult) Then
                    .cell(lRow, lCol).text = sResult
                End If
            
            End If
        End If
    End With
End Sub

Private Sub Grid6_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
lRow = FirstRow
lCol = FirstCol
End Sub

Private Sub Image1_Click()
' delete and reload specifications..


If F_MsgBox.DoShow("Do you want to reload Specifications?") Then
    
    Call CopyFGCodeCertificateGrd2(Grd2, MyCodeID, myCertificate)
    Call CopyFGCodeCertificateGrid2(Grid2, MyCodeID, myCertificate.CalibrationFunction)
    
End If




End Sub

Private Sub ImageTAV_Click(Index As Integer)
Select Case Index
    Case 5
        
End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)

    ExcelApp.Quit
    Set ExcelApp = Nothing
    
Set F_CERTIFICATE = Nothing
End Sub



Private Sub Image3_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub Label2_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub

Private Sub lbExcel_Click()
frExcel_Click
End Sub

Private Sub lbReagent_Click(Index As Integer)
Dim rc As Boolean

    Select Case Index
        Case 50
            rc = True
        Case 51
            rc = False
        Case Else
            Exit Sub
    End Select
    
    lbReagent(50).BackColor = IIf(rc, &H80DF&, &H808080)
    lbReagent(51).BackColor = IIf(Not (rc), &H80DF&, &H808080)
    
    Call ReagentSet(rc)
    
End Sub

Private Sub PicMenu_Click(Index As Integer)



    If IndexFormProcedura = Index Then
    Else
    
        Call SelectProcedura(Index)
    
        
    End If
End Sub


Private Function SelectProcedura(ByVal Index As Integer) As Boolean
Dim rc As Boolean
Dim i As Integer

    'If Index > 3 Then Exit Function
    For i = 0 To PicMenu.Count - 1
        If i = Index Then
            PicMenu(i).BackColor = &H886010
        Else
            PicMenu(i).BackColor = &H644603
            
        End If
    Next
    
  
    
    Select Case Index
        Case 0
        
        Case 1
           
         SetLotCalculation
            
        Case 2
            LoadGraph
            
        Case 3
         
            
    End Select
    
    'Label2(4) = Label2(Index)
    IndexFormProcedura = Index
    PicMain(Index).Visible = True
    PicMain(Index).ZOrder
 
    
   
End Function


Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Integer
For i = 0 To PicMenu.Count - 1
    If i = IndexFormProcedura Then
        ' lascialo stare...
    ElseIf i = Index Then
        PicMenu(i).BackColor = &H886010
    Else
        PicMenu(i).BackColor = &H644603
    End If
Next
End Sub






Private Function GetFormSettingName()
Dim i As Integer
Dim NumReadings As String

   
    CloseSettingDataFile
    
    'For i = 0 To Text1.Count - 1
    '   Text1(i) = GetSettingData(SettingName, "Information QC", "Text1" & i, Text1(i))
    'Next


    '---------------------------------------------------------
    ' numero di letture
    
   
    With MyChemicalQC
        .Lot = GetSettingData(SettingName, "Information QC", "Text10", "")
        .HannaCode.Code = GetSettingData(SettingName, "Information QC", "Text11", "")
        .HannaCode.Description = GetSettingData(SettingName, "Information QC", "Text12", "")
        .Exp = GetSettingData(SettingName, "Information QC", "Text13", "")
        .HannaCode.Line = GetSettingData(SettingName, "Information QC", "Text14", "")
        .HannaCode.Recipe = GetSettingData(SettingName, "Information QC", "Text15", "")
        .date = GetSettingData(SettingName, "Information QC", "Modification Date", "")
        .HannaCode.Decimal = FormatDecimal(GetSettingData(SettingName, "Code Information", "Decimal", 0))
        .HannaCode.MeasurementUnit = GetSettingData(SettingName, "Code Information", "MeasurementUnit", "")
        .ProdFirst = GetSettingData(SettingName, "Information QC", "Text123", "")
        .ProdLast = GetSettingData(SettingName, "Information QC", "Text124", "")
        
        
        
        
       
    End With

    '---------------------------------------------------------

    GetReadingsFormFile

    CloseSettingDataFile
End Function


Private Function CopyLotResultSTD()
Dim k As Integer
Dim t As Integer
Dim Count As Integer
With Grid1
    .AutoRedraw = False
    .Rows = STDCount + 1
    For k = 1 To STDCount
        t = CInt(STD(k, 0))
        If STD(k, 1) = 0 Then
            .Rows = .Rows - 1
            GoTo cont:
        
        End If
        Count = Count + 1
    
       .cell(Count, 0).text = STD(k, 0)
      .cell(Count, 1).text = Format(STD(k, 1), myCertificate.UserDecimal)
      .cell(Count, 2).text = Format(MyChemicalQC.STDtest(t).SelecMean, myCertificate.UserDecimal)
      .cell(Count, 2).FontBold = True
      .cell(Count, 2).ForeColor = vbColorBlueProgram
      .cell(Count, 3).text = k
      
cont:
    Next
    .Column(2).Alignment = cellCenterCenter
    .Column(0).Sort cellAscending
    .Refresh
    .AutoRedraw = True
End With

End Function


Private Function GetReadingsFormFile()
Dim rc              As Boolean
Dim t               As Integer
Dim i               As Integer
Dim j               As Integer
Dim r               As Integer
Dim MaxTest         As Integer
Dim NumReading      As Integer
Dim RedingDiff      As Double
Dim SelTestCount    As Integer
Dim MyCheckString   As String
Dim CheckValue      As String
Dim k               As Integer
Dim strAVG          As String
Dim AverageZero     As Double
Dim CertMaxCount    As Integer

On Error GoTo ERR_READ


    CloseSettingDataFile
   

    MeterNumber = GetSettingData(SettingName, "Information QC", "MeterNumber", 0)
    
    

    If MeterNumber = 0 Then
        PopupMessage 2, "Warning : Please Fill Information QC and Reading QC ...", , True
        CloseSettingDataFile
        Unload Me
        Exit Function
    End If
    
    

    
    CertMaxCount = 0
    Grid3.Rows = 1
    Grid3.AutoRedraw = False
    Dim Count As Integer
    Count = 0
    
    
     
        
    With myCertificate
        .DateAnalisys = GetSettingData(SettingName, "Close QC", "Date", "")
        .UserDecimal = FormatDecimal(GetSettingData(SettingName, "Code Information", "Decimal", 0))
        
    End With
    
    For k = 1 To STDCount
        NumReading = 0
        t = CInt(STD(k, 0))
        MaxTest = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Total Tests", 0)
        
        If MaxTest = 0 Then



                GoTo cont:
               ' Exit Function
        End If
        
        MyChemicalQC.STDtest(t).MaxReadings = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Total Readings", 0)
        MyChemicalQC.STDtest(t).SelReadings = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Selected Readings", 0)
        MyChemicalQC.STDtest(t).NumTest = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Total Tests", 0)
    
        MyChemicalQC.STDtest(t).TotalMean = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Total Average", 0)
       ' MyChemicalQC.STDtest(t).SelecMean = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Selected Average", 0)
        
        Count = Count + 1



        With myCertificate
        
               
                
            .STD(Count).AverageResult = 0
            
            
            
        End With
         
        SelTestCount = 0
        AverageZero = 0
       
        
        For i = 1 To MaxTest
            
            
            strAVG = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Test " & i & " ReadingAvg", "NULL")
            
            If strAVG = "NULL" Then
                GoTo cont
            End If
            
            
            
                With MyChemicalQC.STDtest(t)

                        NumReading = NumReading + 1

                        .Readings(NumReading).STDAbs = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Test " & i & " ABS", 0)
                        .Readings(NumReading).STDValueAvg = CDbl(strAVG)
                        
                        If .Readings(NumReading).STDAbs = "" Then
                            GoTo cont
                        End If
                        
                        
                        myCertificate.STD(Count).AverageResult = myCertificate.STD(Count).AverageResult + .Readings(NumReading).STDValueAvg
                        

                            SelTestCount = SelTestCount + 1
                            
                            If STD(k, 1) <> 0 Then
                            
                                CertMaxCount = CertMaxCount + 1
                            
                                With myCertificate.LotCalculation
                                
                                    ReDim Preserve .STDValue(CertMaxCount)
                                    ReDim Preserve .STDLotResult(CertMaxCount)
                                    ReDim Preserve .STDYcalc(CertMaxCount)
                                    ReDim Preserve .Uncertainty(CertMaxCount)
                                    ReDim Preserve .UncertPerc(CertMaxCount)
                                   
                                    
                                    .STDValue(CertMaxCount) = STD(k, 1)
                                    
                                    
                                    .STDLotResult(CertMaxCount) = MyChemicalQC.STDtest(t).Readings(NumReading).STDAbs
                                    
                                    
                                    .MaxValue = IIf(.STDValue(CertMaxCount) > .MaxValue, .STDValue(CertMaxCount), .MaxValue)
                                    .MinValue = IIf(.STDValue(CertMaxCount) < .MinValue, .STDValue(CertMaxCount), .MinValue)
                                    If .MinValue = 0 Then .MinValue = .STDValue(CertMaxCount)

                                    
                                    Grid3.AddItem "", False
                                    Grid3.cell(Grid3.Rows - 1, 0).text = CertMaxCount
                                    Grid3.cell(Grid3.Rows - 1, 1).text = STD(k, 0)
                                    Grid3.cell(Grid3.Rows - 1, 2).text = STD(k, 1)
                                    Grid3.cell(Grid3.Rows - 1, 3).text = .STDLotResult(CertMaxCount)
                                    
                                   .MaxCount = CertMaxCount
                                End With
                            Else
                                ' STD 0 bisogna calcolare la media
                                'myCertificate.STD(Count).AverageResult = myCertificate.STD(Count).AverageResult + MyChemicalQC.STDtest(t).Readings(NumReading).STDAbs
                                
                            End If
                End With
cont:
        Next
        
        
        myCertificate.STD(Count).AverageResult = myCertificate.STD(Count).AverageResult / NumReading
        MyChemicalQC.STDtest(t).SelecMean = myCertificate.STD(Count).AverageResult
        
        If STD(k, 1) = 0 Then
               myCertificate.CalibrationFunction.ReagentBlank.LotValue = Format(MyChemicalQC.STDtest(t).SelecMean, "0.000")  'MyChemicalQC.STDtest(t).SelecMean
        End If
         
    Next
ERR_END:
    On Error GoTo 0
    Debug.Print Grid3.Rows
    Grid3.ReadOnly = True
    Grid3.Refresh
    Grid3.AutoRedraw = True
    
    
    Exit Function
ERR_READ:
    MsgBox Err.Description
    Resume Next


End Function






Private Function GetReadingsFormFile_OLD()
Dim rc              As Boolean
Dim t               As Integer
Dim i               As Integer
Dim j               As Integer
Dim r               As Integer
Dim MaxTest         As Integer
Dim NumReading      As Integer
Dim RedingDiff      As Double
Dim SelTestCount    As Integer
Dim MyCheckString   As String
Dim CheckValue      As String
Dim k               As Integer
Dim strAVG          As String

Dim CertMaxCount    As Integer

On Error GoTo ERR_READ



'Public Type CertResult
'    MaxCount            As Integer
'    STDValue()          As Double
'    STDLotResult()      As Double
'    STDYcalc()          As Double
'    a                   As Double
'    b                   As Double
    
'End Type

    CloseSettingDataFile
   

    MeterNumber = GetSettingData(SettingName, "Information QC", "MeterNumber", 0)
    
    

    If MeterNumber = 0 Then
        PopupMessage 2, "Warning : Please Fill Information QC and Reading QC ...", , True
        CloseSettingDataFile
        Unload Me
        Exit Function
    End If
    
    

    
    CertMaxCount = 0
    Grid3.Rows = 1
    Grid3.AutoRedraw = False
    Dim Count As Integer
    Count = 0
    
    
     
        
    With myCertificate
        .DateAnalisys = GetSettingData(SettingName, "Close QC", "Date", "")
        .UserDecimal = FormatDecimal(GetSettingData(SettingName, "Code Information", "Decimal", 0))
        
    End With
    
    For k = 1 To STDCount
        NumReading = 0
        t = CInt(STD(k, 0))
        MaxTest = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Total Tests", 0)
        
        If MaxTest = 0 Then



                GoTo cont:
               ' Exit Function
        End If
        
        MyChemicalQC.STDtest(t).MaxReadings = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Total Readings", 0)
        MyChemicalQC.STDtest(t).SelReadings = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Selected Readings", 0)
        MyChemicalQC.STDtest(t).NumTest = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Total Tests", 0)
       ' MyChemicalQC.STDtest(t).SelTest = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Selected Tests", 0)
        MyChemicalQC.STDtest(t).TotalMean = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Total Average", 0)
        MyChemicalQC.STDtest(t).SelecMean = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Selected Average", 0)
        
        Count = Count + 1



        With myCertificate
        
               
                

            .STD(Count).AverageResult = MyChemicalQC.STDtest(t).SelecMean
            
            If STD(k, 1) = 0 Then
               .CalibrationFunction.ReagentBlank.LotValue = Format(MyChemicalQC.STDtest(t).SelecMean, .UserDecimal)
            End If
         
        End With
         
        SelTestCount = 0
      
       
        
        For i = 1 To MaxTest
            
            
            strAVG = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Test " & i & " ReadingAvg", "NULL")
            
            If strAVG = "NULL" Then
                GoTo contTest
            End If
            
            For j = 1 To MeterNumber
                With MyChemicalQC.STDtest(t)
                    
                    CheckValue = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Test " & i & " Meter " & j & " Value", "")
                    
                    If CheckValue <> "" Then
                        NumReading = NumReading + 1
                        
                        .Readings(NumReading).Value = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Test " & i & " Meter " & j & " Value", "")
                        .Readings(NumReading).Meter = j
                        .Readings(NumReading).Test = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Test " & i & " Real Test", "")
                        .Readings(NumReading).bSelectedValue = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Test " & i & " Meter " & j & " Selected", "TRUE")
                        
                        
                        
                         
                        ' SaveSettingData SettingName, "Graph QC", "Standard " & t & " Test " & i & " ABS", .Cell(t, 15).Text
    
                        
                        '.Readings(NumReading).STDAbs
                        
                       ' .Readings(NumReading).STDValueAvg = CDbl(strAVG)
                        
                       
                        
                        If .Readings(NumReading).bSelectedValue = True Then
                            
                            SelTestCount = SelTestCount + 1
                            
                            If STD(k, 1) <> 0 Then
                            
                                CertMaxCount = CertMaxCount + 1
                            
                                With myCertificate.LotCalculation
                                
                                    ReDim Preserve .STDValue(CertMaxCount)
                                    
                                    
                                    ReDim Preserve .STDLotResult(CertMaxCount)
                                    ReDim Preserve .STDYcalc(CertMaxCount)
                                    ReDim Preserve .Uncertainty(CertMaxCount)
                                    ReDim Preserve .UncertPerc(CertMaxCount)
                                   
                                    
                                    .STDValue(CertMaxCount) = STD(k, 1)
                                    .STDLotResult(CertMaxCount) = MyChemicalQC.STDtest(t).Readings(NumReading).Value
                                    
                                    
                                    .MaxValue = IIf(.STDValue(CertMaxCount) > .MaxValue, .STDValue(CertMaxCount), .MaxValue)
                                    .MinValue = IIf(.STDValue(CertMaxCount) < .MinValue, .STDValue(CertMaxCount), .MinValue)
                                    If .MinValue = 0 Then .MinValue = .STDValue(CertMaxCount)

                                    
                                    Grid3.AddItem "", False
                                    Grid3.cell(Grid3.Rows - 1, 0).text = CertMaxCount
                                    Grid3.cell(Grid3.Rows - 1, 1).text = STD(k, 0)
                                    Grid3.cell(Grid3.Rows - 1, 2).text = STD(k, 1)
                                    Grid3.cell(Grid3.Rows - 1, 3).text = .STDLotResult(CertMaxCount)
                                    
                                   .MaxCount = CertMaxCount
                                End With
                            End If
                        End If
                    End If
                    
                End With
contTest:
            Next
cont:
        Next

    Next
ERR_END:
    On Error GoTo 0
    Debug.Print Grid3.Rows
    Grid3.ReadOnly = True
    Grid3.Refresh
    Grid3.AutoRedraw = True
    
    
    Exit Function
ERR_READ:
    MsgBox Err.Description
    Resume Next


End Function

Private Function GetSTDValueForGraph() As Boolean
Dim t As Integer
Dim k As Integer
Dim rc As Boolean
Dim Tolerance As Double
Dim CheckIt As String

CloseSettingDataFile


   If FileExists(USER_TEMP_PATH & SettingName) Then
   ElseIf FileExists(USER_DATA_PATH & SettingName) Then
       ' PopupMessage 2, "Lot : " & Text1(0) & vbCrLf & "Code : " & Text1(1) & vbCrLf & "Is Closed..."
        USER_PATH = USER_DATA_PATH
   Else
        GetSTDValueForGraph = False
   End If
   
   
   
   
        STDCount = GetSettingData(SettingName, "Graph QC", "STDCount", 0)
        Virgola = GetSettingData(SettingName, "Code Information", "Decimal", 0)
        UserDecimal = FormatDecimal(GetSettingData(SettingName, "Code Information", "Decimal", 0))
        MeasurementUnit = GetSettingData(SettingName, "Code Information", "MeasurementUnit", "")
        
    If STDCount = 0 Then
        GetSTDValueForGraph = False
        Exit Function
    End If
        

        
        
        With myCertificate
        
            ReDim .STD(STDCount)
            ReDim STD(STDCount, 3) As String

                For k = 1 To STDCount
                    CheckIt = GetSettingData(SettingName, "Graph QC", "STDValue" & k, "")
                    
                    If CheckIt = "" Or CheckIt = 0 Then
                    ' GoTo cont
                    End If
                    
                    t = t + 1
                    
                    STD(t, 0) = GetSettingData(SettingName, "Graph QC", "STDNumber" & k, "")
                    STD(t, 1) = GetSettingData(SettingName, "Graph QC", "STDValue" & k, "")
                    STD(t, 2) = GetSettingData(SettingName, "Graph QC", "STDMin" & k, "")
                    STD(t, 3) = GetSettingData(SettingName, "Graph QC", "STDMax" & k, "")
                    
                    If InStr(STD(t, 1), "/") Or STD(t, 1) = "" Then
                        STD(t, 1) = 0
                    End If
                    If InStr(STD(t, 2), "/") Or STD(t, 2) = "" Then
                        STD(t, 2) = 0
                    End If
                    If InStr(STD(t, 3), "/") Or STD(t, 3) = "" Then
                        STD(t, 3) = 0
                    End If
                    CloseSettingDataFile
                    
                    Tolerance = CDbl(STD(t, 3)) - CDbl(STD(t, 1))
                   
                    With MyChemicalQC.DataControllo(k)
                        .s = CDbl(STD(t, 2)) - CDbl(STD(t, 1))
                        .s2 = CDbl(STD(t, 3)) - CDbl(STD(t, 1))
                        .STDRef = CDbl(STD(t, 1))
                        .STDMin = STD(t, 2)
                        .STDMax = STD(t, 3)
                        .STDNumber = STD(t, 0)
                    End With
                    
                    .STD(t).STDValue = STD(t, 1)
                   
cont:
                Next
                
                STDCount = t
        End With
        
       
    CloseSettingDataFile
    
    GetSTDValueForGraph = True

End Function








































'--------------------------------------------
'
'
'
'               GRAFICO
'
'
'
'---------------------------------------------




Private Sub chChart_AxisLabelActivated(axisID As Integer, AxisIndex As Integer, labelSetIndex As Integer, LabelIndex As Integer, MouseFlags As Integer, Cancel As Integer)
    'Activate AxisLabel and put values in textbox
    'put new scale and apply to chart
    'If click at Y or Y2 Axis
    If axisID = VtChAxisIdY Or axisID = VtChAxisIdY2 Then
        With chChart.Plot.Axis(VtChAxisIdY).ValueScale
          '  txtValue(0).Text = .Maximum
          '  txtValue(1).Text = .Minimum
          '  txtValue(2).Text = .MajorDivision
          '  txtValue(3).Text = .MinorDivision
        End With
       ' cmdScale.Enabled = True
    End If
End Sub

'Change Axis TItle Text
Private Sub chChart_AxisTitleActivated(axisID As Integer, AxisIndex As Integer, MouseFlags As Integer, Cancel As Integer)
    Dim InputBoxTitle As String
    Dim InputText As String
    Dim CurrentText As String
    
    If axisID = VtChAxisIdX Then
        InputBoxTitle = "Asse X"
    ElseIf axisID = VtChAxisIdY Then
        InputBoxTitle = "Asse Y"
    End If
    
    'Get current text
    CurrentText = chChart.Plot.Axis(axisID).AxisTitle.text
    'Show inputbox
    InputText = CurrentText
    If F_InputBox.DoShow("Inserire il titolo del Grafico :", "Grafico", , , , InputText) Then
        If InputText <> "" Then
            chChart.Plot.Axis(axisID).AxisTitle.text = InputText
            If axisID = VtChAxisIdX Then
                SaveSetting App.Title, "Grafico", "X Text", InputText
            Else
                SaveSetting App.Title, "Grafico", "Y Text", InputText
            End If
        End If
    End If
End Sub

'Show value in ToolTipText
Private Sub chChart_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim Part As Integer, Series As Integer, DataPoint As Integer
    Dim Index3 As Integer, Index4 As Integer
    Dim oValue As Double, NullFlag As Integer

    With chChart
    
         'Already set at design-time
'        .AllowDithering = False
'        .AllowSelections = True
'        .AutoIncrement = False
'        .AllowDynamicRotation = False
'        .AllowSeriesSelection = False
        
        .TwipsToChartPart x, Y, Part, Series, DataPoint, Index3, Index4
    
        'lblXValue.Caption = X
        'lblYValue.Caption = Y
        'lblPart.Caption = Part
        'lblSerie.Caption = Series
        'lblData.Caption = DataPoint
        
        'Show value in ToolTipText when select point
        If Part = VtChPartTypeChart Then
            .ToolTipText = "Chart Area"
        ElseIf Part = VtChPartTypeTitle Then
            .ToolTipText = "Chart Title"
        ElseIf Part = VtChPartTypeFootnote Then
            .ToolTipText = "Footnote"
        ElseIf Part = VtChPartTypeLegend Then
            .ToolTipText = "Legend"
        ElseIf Part = VtChPartTypePlot Then
            .ToolTipText = "Plot Area"
        ElseIf Part = VtChPartTypePoint Or Part = VtChPartTypePointLabel Then
            .DataGrid.GetData DataPoint, Series, oValue, NullFlag
            .ToolTipText = "Value := " & FormatNumber(oValue, 2)
        ElseIf Part = VtChPartTypeAxis Then
            .ToolTipText = "Plot Axis"
        ElseIf Part = VtChPartTypeAxisLabel Then
            .ToolTipText = "Axis Label"
        ElseIf Part = VtChPartTypeAxisTitle Then
            .ToolTipText = "Axis Title"
        Else
            .ToolTipText = ""
        End If
    End With
End Sub

'Change wall color
Private Sub chChart_PlotActivated(MouseFlags As Integer, Cancel As Integer)

    
    With dlgOption
        .CancelError = False
        .flags = cdlCCRGBInit
        .ShowColor
        oRed = RedFromRGB(.Color)
        oGreen = GreenFromRGB(.Color)
        oBlue = BlueFromRGB(.Color)
        
        SaveSetting App.Title, "Grafico", "Backgrond oRed", oRed
        SaveSetting App.Title, "Grafico", "Backgrond oGreen", oGreen
        SaveSetting App.Title, "Grafico", "Backgrond oBlue", oBlue
        
        chChart.Plot.Wall.Brush.FillColor.Set oRed, oGreen, oBlue
    End With
End Sub

Private Sub chChart_PointActivated(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
    Call chChart_SeriesActivated(Series, MouseFlags, Cancel)
End Sub

Private Sub chChart_PointLabelActivated(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
    Call chChart_SeriesActivated(Series, MouseFlags, Cancel)
End Sub

'Change serie color
Private Sub chChart_SeriesActivated(Series As Integer, MouseFlags As Integer, Cancel As Integer)

    With dlgOption
        .CancelError = False
        .flags = cdlCCRGBInit
        .ShowColor
        oRed = RedFromRGB(.Color)
        oGreen = GreenFromRGB(.Color)
        oBlue = BlueFromRGB(.Color)
        
            SaveSetting App.Title, "Grafico", "S" & Series & " oRed", oRed
            SaveSetting App.Title, "Grafico", "S" & Series & " oGreen", oGreen
            SaveSetting App.Title, "Grafico", "S" & Series & " oBlue", oBlue
            

        With chChart.Plot.SeriesCollection(Series).DataPoints(-1)
            'for 2d bar
            .Brush.FillColor.Set oRed, oGreen, oBlue
            'data label
            .DataPointLabel.VtFont.VtColor.Set oRed, oGreen, oBlue
            'for 2d line
            .Marker.Pen.VtColor.Set oRed, oGreen, oBlue
        End With
    End With
End Sub

'Change Title Text
Private Sub chChart_TitleActivated(MouseFlags As Integer, Cancel As Integer)
    Dim InputText As String
    Dim CurrentText As String
        
    'Get current text
    CurrentText = chChart.TitleText
    'Show inputbox
    InputText = CurrentText
    If F_InputBox.DoShow("Inserire il campo Titolo", "Grafico", , , , InputText) Then
        
        If InputText <> "" Then
            chChart.TitleText = InputText
            SaveSetting App.Title, "Grafico", "Title Text", InputText
            
        End If
    End If
End Sub



Private Sub SaveGraphJPG()
    Dim rv As Long
    Dim FileName As String
    On Error GoTo ErrHandler
    
    If Me.Visible = False Then Exit Sub
    

        FileName = USER_TEMP_PATH & "\GraficoErrore" & ".jpg"
        

            Screen.MousePointer = vbHourglass
            With picGraph
                .Height = chChart.Height
                .Width = chChart.Width
                .AutoRedraw = True
                rv = SendMessage(chChart.hWnd, WM_PAINT, .hdc, 0)
                .Picture = .Image
                .AutoRedraw = False
            End With
         
            ' Sent the picture to the clipboard.
            Clipboard.Clear
            Clipboard.SetData picGraph.Picture
            If picGraph.Picture = 0 Then Exit Sub
            SavePicture picGraph.Picture, FileName
            Debug.Print FileName
            Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    MsgBox Err.Description
    Resume Next
End Sub

Private Sub LoadGraph()
    Dim i As Integer
    Dim k As Integer
    Dim STDMaxCount As Integer
    Dim Graph() As Variant
   
    Dim NumLinee As Integer
    
    Dim TargetValue()   As Double
    Dim MeanValue()     As Double
    Dim Llimit()        As Double
    Dim Ulimit()        As Double
    
On Error GoTo ERR_GRAPH
    
    NumLinee = 4
    STDMaxCount = 0
    
     For k = 1 To STDCount
        If STD(k, 1) = 0 Then
            GoTo cont:
        End If
        STDMaxCount = STDMaxCount + 1
cont:
    Next
    
    
    If STDMaxCount = 0 Then
        ReDim Graph(1 To 6, 1 To 5) As Variant
        GoSub STDValue
        GoTo Grafico
    End If
    
    ReDim Graph(1 To STDMaxCount, 1 To NumLinee) As Variant
    ReDim TargetValue(STDMaxCount) As Double
    ReDim MeanValue(STDMaxCount) As Double
    ReDim Llimit(STDMaxCount) As Double
    ReDim Ulimit(STDMaxCount) As Double
    
    Dim Division As Long
    
    For i = 1 To STDMaxCount
        
        
        Llimit(i) = myCertificate.GraphCert.LplimGrph(i)
        Ulimit(i) = myCertificate.GraphCert.UplimGrph(i)
        TargetValue(i) = myCertificate.STD(i).STDValue
        MeanValue(i) = myCertificate.STD(i).AverageResult
      
        Graph(i, 1) = CStr(TargetValue(i)) & " ppm"
        Graph(i, 2) = Llimit(i)
        Graph(i, 3) = Ulimit(i)
        Graph(i, 4) = MeanValue(i)
    
    Next

Grafico:

    chChart.ChartType = VtChChartType2dLine ' VtChChartType2dBar

     chChart.ChartData = Graph

    chChart.ShowLegend = True
    chChart.Legend.Location.LocationType = VtChLocationTypeTop
    
    
    'Title
    Dim sString As String
    sString = GetSetting(App.Title, "Grafico", "Title Text", "STD Chart")

    chChart.Title.text = sString '"Product Quality Chart"
    
    'Axis Title
    sString = GetSetting(App.Title, "Grafico", "X Text", "Target Value (ppm)")
    chChart.Plot.Axis(VtChAxisIdX).AxisTitle.text = sString '"Target Value (ppm)"
    
    sString = GetSetting(App.Title, "Grafico", "Y Text", "Result (ppm)")
    chChart.Plot.Axis(VtChAxisIdY).AxisTitle.text = sString '"Result (ppm)"
    
    
    'If wanna hide Y2 axis
    chChart.Plot.Axis(VtChAxisIdY2).AxisScale.Hide = False ' True
    
    'Legend
   ' chChart.Plot.SeriesCollection(1).LegendText = "Product ID := AAA"
    'chChart.Plot.SeriesCollection(2).LegendText = "Product ID := BBB"
    'chChart.Plot.SeriesCollection(3).LegendText = "Product ID := CCC"
   ' chChart.Legend.Location.LocationType = VtChLocationTypeTopRight
    
    'Wall
    chChart.Plot.Wall.Brush.Style = VtBrushStyleSolid

    oRed = GetSetting(App.Title, "Grafico", "Backgrond oRed", 255)
    oGreen = GetSetting(App.Title, "Grafico", "Backgrond oGreen", 255)
    oBlue = GetSetting(App.Title, "Grafico", "Backgrond oBlue", 255)
        
        
    chChart.Plot.Wall.Brush.FillColor.Set oRed, oGreen, oBlue
   
    
    With chChart.Plot.Axis(VtChAxisIdX)
           .AxisScale.Hide = False
           .AxisGrid.MajorPen.Style = VtPenStyleDotted
                With .CategoryScale
                    .Auto = False
                    .DivisionsPerLabel = 1
                    .DivisionsPerTick = 1
                    .LabelTick = True
                End With
                With .Labels(1)
                    .Auto = False
                    .TextLayout.Orientation = VtOrientationHorizontal
                    .VtFont.VtColor.Set 0, 0, 0
                End With
                With .AxisGrid.MajorPen
                    .Style = VtPenStyleDitted
                    .VtColor.Set 0, 0, 0
                End With
    End With
    
    With chChart.Plot.Axis(VtChAxisIdY)
           .AxisScale.Hide = False
           .AxisGrid.MajorPen.Style = VtPenStyleDotted
           
           '.AxisGrid.MajorPen.VtColor.Set 255, 255, 255
           '.AxisGrid.MinorPen.VtColor.Set 255, 255, 255
    End With
    
    
    With chChart.Plot.Axis(VtChAxisIdY).ValueScale
        .Auto = False
        
        .Maximum = Graph(STDMaxCount, 2) * 1.2
        .Minimum = 0
        .MajorDivision = STDMaxCount - 1
        .MinorDivision = 1
    
      
    End With
 
        With chChart.Plot
            .AngleUnit = VtAngleUnitsGrads
            
            oRed = GetSetting(App.Title, "Grafico", "S1 oRed", 255)
            oGreen = GetSetting(App.Title, "Grafico", "S1 oGreen", 0)
            oBlue = GetSetting(App.Title, "Grafico", "S1 oBlue", 0)
            

            With .SeriesCollection(1)
            
                .LegendText = "Lower Limit"
                With .Pen
                    .Width = ScaleX(3, vbPixels, vbTwips)
                    .VtColor.Set oRed, oGreen, oBlue
                End With
            End With
            
            
            oRed = GetSetting(App.Title, "Grafico", "S2 oRed", 255)
            oGreen = GetSetting(App.Title, "Grafico", "S2 oGreen", 0)
            oBlue = GetSetting(App.Title, "Grafico", "S2 oBlue", 0)
            
            With .SeriesCollection(2)
            
                .LegendText = "Upper Limit"
                With .Pen
                    .Width = ScaleX(3, vbPixels, vbTwips)
                    .VtColor.Set oRed, oGreen, oBlue
                End With
            End With
            
            oRed = GetSetting(App.Title, "Grafico", "S3 oRed", 255)
            oGreen = GetSetting(App.Title, "Grafico", "S3 oGreen", 0)
            oBlue = GetSetting(App.Title, "Grafico", "S3 oBlue", 0)
            
            With .SeriesCollection(3)
            
                .LegendText = "STD Mean Value"
                With .Pen
                    .Width = ScaleX(3, vbPixels, vbTwips)
                    .VtColor.Set oRed, oGreen, oBlue
                End With
            End With
                        
            
            
            
        End With


    'Serie
   ' For i = 1 To 4
    i = 3
    
        With chChart.Plot.SeriesCollection(i).DataPoints(-1)
        
        
        
            '.DataPointLabel.LocationType = VtChLabelLocationTypeBelowPoint ' VtChLabelLocationTypeCenter
            .DataPointLabel.ValueFormat = "0.00"
            .DataPointLabel.VtFont.Name = "CALIBRI"
            .DataPointLabel.VtFont.Size = 10
            .DataPointLabel.VtFont.Style = VtFontStyleOutline 'Regular

            'Label color by serie
            If i = 1 Then
                .DataPointLabel.VtFont.VtColor.Set 255, 64, 64
            ElseIf i = 2 Then
                .DataPointLabel.VtFont.VtColor.Set 128, 128, 255
            ElseIf i = 3 Then
                .DataPointLabel.VtFont.VtColor.Set 64, 192, 192
            End If
        
            'Show Marker
           ' chChart.Plot.SeriesCollection(i).SeriesMarker.Show = True
        End With
        
        
        ' salvo subito!!!
        
        
         SaveGraphJPG
        
        
        
        Exit Sub
        
ERR_END:
        On Error GoTo 0
        Exit Sub
ERR_GRAPH:

        MsgBox Err.Description
        Resume Next
        
        Exit Sub
        
STDValue:
    Graph(1, 1) = "500g"
    Graph(2, 1) = "1.000g"
    Graph(3, 1) = "2.000g"
    Graph(4, 1) = "5.000g"
    Graph(5, 1) = "10.000g"
    Graph(6, 1) = "20.000g"
    
    ' tolleranza fissa
    
    'For i = 0 To 6000
    '    Graph1(i, 1) = 45
    '   Graph1(i, 2) = -45
    'Next
    
    ' tolleranza EMT
    

    Graph(1, 2) = 15
    Graph(2, 2) = 30
    Graph(3, 2) = 30
    Graph(4, 2) = 45
    Graph(5, 2) = 45
    Graph(6, 2) = 45
    
    
   
    Graph(1, 3) = -15
    Graph(2, 3) = -30
    Graph(3, 3) = -30
    Graph(4, 3) = -45
    Graph(5, 3) = -45
    Graph(6, 3) = -45
    
    
    ' Errore Corretto

  
    Graph(1, 4) = 8
    Graph(2, 4) = 12
    Graph(3, 4) = 13
    Graph(4, 4) = 18
    Graph(5, 4) = 18
    Graph(6, 4) = 20
   
    Graph(1, 5) = -10
    Graph(2, 5) = -25
    Graph(3, 5) = -28
    Graph(4, 5) = -35
    Graph(5, 5) = -38
    Graph(6, 5) = -40
    
    STDMaxCount = 6
    NumLinee = 5
   Return
        
        
  '  Next i
End Sub



Private Function RedFromRGB(ByVal oRGB As Long) As Integer
   RedFromRGB = &HFF& And oRGB
End Function

Private Function GreenFromRGB(ByVal oRGB As Long) As Integer
   GreenFromRGB = (&HFF00& And oRGB) \ 256
End Function

Private Function BlueFromRGB(ByVal oRGB As Long) As Integer
   BlueFromRGB = (&HFF0000 And oRGB) \ 65536
End Function



















'---------------------------------------------------------------------------------
'
'       MyCertificate
'
'---------------------------------------------------------------------------------


Private Function SetLotCalculation()
Dim gdl As Integer
Dim MaxCount As Integer
Dim i As Integer


NumLinearità = 9

Set ExcelApp = CreateObject("Excel.Application")

    
    With myCertificate.LotCalculation
        MaxCount = .MaxCount

        Grid3.AutoRedraw = False
        If MaxCount = 0 Then Exit Function
        
        
         myCertificate.CalibrationFunction.gdl.LotValue = IIf(MaxCount > 0, MaxCount - 2, 0)
         
         gdl = myCertificate.CalibrationFunction.gdl.LotValue

        .df = gdl
        
        
       ' .a = CalculateIntercept(.STDLotResult, .STDValue)
       ' .b = CalculateSlope(.STDValue, .STDLotResult)
        
        .a = ExcelApp.WorksheetFunction.Intercept(.STDLotResult, .STDValue)
        .b = ExcelApp.WorksheetFunction.Slope(.STDLotResult, .STDValue)
        
      
        .n = MaxCount
        .MedValue = 0
        For i = 1 To MaxCount
            
            .STDYcalc(i) = .a + .b * .STDValue(i)
            .MedValue = .MedValue + .STDValue(i)
            
          If Grid3.Rows > i Then
            Grid3.cell(i, 4).text = FormatNumber(.STDYcalc(i), 3)
            Grid3.cell(i, 4).BackColor = vbColorAzzurrino
          End If
        Next
        .MedValue = .MedValue / MaxCount
        
        
        
        .RSS = FormatNumber(ExcelApp.WorksheetFunction.SumXMY2(.STDLotResult, .STDYcalc), 6)
        .TSS = FormatNumber(ExcelApp.WorksheetFunction.devSq(.STDLotResult), 6)
        
        'SIGN(b)*(1-RSS/TSS)^0,5
        .r = Abs(.b) * (1 - (.RSS / .TSS))
        
        
        Grid3.Column(0).Width = 15
        Grid3.RowHeight(0) = 80
        For i = 1 To Grid3.Cols - 1
             Grid3.cell(0, i).WrapText = True
             Grid3.Column(i).Width = 70
             Grid3.Column(i).Alignment = cellCenterCenter
        Next
         
   
       ' .MedValue = (.MinValue + .MaxValue) / 2
      '  .sy = CalculateSTEYX(.STDLotResult, .STDValue)
        
        .sy = ExcelApp.WorksheetFunction.SteYX(.STDLotResult, .STDValue)
        
       ' .ssx = CalculateDevSq(.STDValue)
         .ssx = ExcelApp.WorksheetFunction.devSq(.STDValue)
   
        .tval = FormatNumber(ExcelApp.WorksheetFunction.T_Inv_2T(0.0455, gdl), 8)
        
        .MethodStDeviation = .sy / .b
        .MethodVariation = FormatNumber((.MethodStDeviation / .MedValue) * 100, 4)
        
        
        
        '"Max Uncertainty sample "
        'sy/b*(1/P27+1/n+(R27-AVERAGE(Y))^2/(b^2*SSx))^0,5
        Dim Replicat    As Integer
        Dim Diff        As Double
        Dim Denom       As Double
        Dim AverageY    As Double
        
        AverageY = ExcelApp.WorksheetFunction.Average(.STDLotResult)
        Replicat = 1
        
        For i = 1 To MaxCount
        
            Diff = .STDYcalc(i) - AverageY
            Denom = (.b ^ 2) * .ssx
            .Uncertainty(i) = (.sy / .b) * Sqr((1 / Replicat) + (1 / .n) + Diff ^ 2 / Denom)
            .UncertPerc(i) = FormatNumber(.Uncertainty(i) / .STDValue(i), 4) * 100
          
            Grid3.cell(i, 5).text = FormatNumber(.Uncertainty(i), 3)
            Grid3.cell(i, 6).text = FormatNumber(.UncertPerc(i), 2)
            Grid3.cell(i, 7).text = FormatNumber(.Uncertainty(i) * .tval, 3)
            Grid3.cell(i, 7).BackColor = vbColorAzzurrino
            
        Next
        
        
        
        
        .ConfidenceInterval = .tval * GetMax(.Uncertainty)
        
        

        
        
        Grid4.cell(1, 1).text = FormatNumber(.a, 3)
        Grid4.cell(1, 2).text = FormatNumber(.b, 3)
        Grid4.cell(1, 3).text = .n
        Grid4.cell(1, 4).text = FormatNumber(.tval, 2)
        Grid4.cell(1, 5).text = FormatNumber(.MedValue, 3)
        Grid4.cell(1, 6).text = FormatNumber(.sy, 4)
        Grid4.cell(1, 7).text = FormatNumber(.ssx, 1)
        Grid4.cell(1, 8).text = FormatNumber(.MethodStDeviation, 3)
        Grid4.cell(1, 9).text = FormatNumber(.MethodVariation, 3)
        Grid4.cell(1, 10).text = FormatNumber(.ConfidenceInterval, 3)
        
         
        Grid4.Column(0).Width = 0
        Grid4.RowHeight(0) = 80
        For i = 1 To Grid4.Cols - 1
            Grid4.Column(i).Width = 100
        Next
         
         
         
         
        
    'MinValue            As Double
    'MaxValue            As Double
    'MedValue            As Double
    
    'a                   As Double
    'b                   As Double
    
    'r                   As Double
    'sy                  As Double
    'MethodStDeviation   As Double
    'MethodVariation     As Double
    
    
    

    Grid3.Refresh
    Grid3.AutoRedraw = True
    
    
    ' Get Max StdDeviation
    Call GetMaxStdDeviation
    
    
    End With
    With myCertificate.CalibrationFunction
    

       '
       ' .Cell(1, 0).Text = "n"
       ' .Cell(2, 0).Text = "Standard Deviation [mg/L NH4+]"
       ' .Cell(3, 0).Text = "Method variation coefficient [%]"
       ' .Cell(4, 0).Text = "Confidence interval (95%) [mg/L NH4+]"

       ' .Cell(5, 0).Text = "Slope"
       ' .Cell(6, 0).Text = "Ordinate intersect ppm"
       ' .Cell(7, 0).Text = "Blank Value [Absorbance]"
       
       
        '.Cell(2, 0).Text = "Standard Deviation [mg/L NH4+]"
        '.Cell(4, 0).Text = "Confidence interval (95%) [mg/L NH4+]"
        Dim iDecimal As String
       
      
        .Slope.LotValue = FormatNumber(myCertificate.LotCalculation.r, 4)
        .Intersect.LotValue = FormatNumber(myCertificate.LotCalculation.a, 4)
        .Variation.LotValue = FormatNumber(myCertificate.LotCalculation.MethodVariation, 4)
        .Confidence.LotValue = FormatNumber(myCertificate.LotCalculation.ConfidenceInterval, 4)
        
        .gdl.Passed = IIf(.gdl.TargetValue <= .gdl.LotValue, True, False)
        .Slope.Passed = IIf(.Slope.TargetValue <= .Slope.LotValue, True, False)
        .ReagentBlank.Passed = IIf(.ReagentBlank.TargetValue >= .ReagentBlank.LotValue, True, False)
        .Variation.Passed = IIf(.Variation.TargetValue >= .Variation.LotValue, True, False)
        .Confidence.Passed = IIf(.Confidence.TargetValue >= .Confidence.LotValue, True, False)
        .StdDeviation.Passed = IIf(.StdDeviation.TargetValue >= .StdDeviation.LotValue, True, False)
        
        .Intersect.Passed = IIf(.Intersect.TargetValue >= .Intersect.LotValue, True, False)
        
        
        
        Grid2.cell(1, 2).text = .gdl.LotValue
        Grid2.cell(2, 2).text = .StdDeviation.LotValue
        Grid2.cell(3, 2).text = .Variation.LotValue
        Grid2.cell(4, 2).text = .Confidence.LotValue
        
        Grid2.cell(5, 2).text = .Slope.LotValue
        Grid2.cell(6, 2).text = .Intersect.LotValue
        Grid2.cell(7, 2).text = .ReagentBlank.LotValue

        
        Grid2.Column(3).CellType = cellCheckBox
        
        Grid2.cell(1, 3).text = .gdl.Passed
        Grid2.cell(2, 3).text = .StdDeviation.Passed
        Grid2.cell(3, 3).text = .Variation.Passed
        Grid2.cell(4, 3).text = .Confidence.Passed
        Grid2.cell(5, 3).text = .Slope.Passed
        Grid2.cell(6, 3).text = .Intersect.Passed
        Grid2.cell(7, 3).text = .ReagentBlank.Passed

        Grid2.Column(1).AutoFit
        Grid2.Column(2).AutoFit
        
        For i = 1 To Grid2.Rows - 1
            
            Grid2.cell(i, 2).FontBold = True
            If Grid2.cell(i, 3).text = True Then
                Grid2.cell(i, 2).BackColor = vbColorAzzurrino
            Else
                Grid2.cell(i, 2).BackColor = vbColorRosaTabella
            End If
        Next
        
    
    End With
     
    SetCalibrationUncertainty


End Function




Private Function SetCalibrationUncertainty()
Dim Maxx As Double
Dim Minx As Double
Dim i As Integer
Dim a As Double
Dim b As Double
Dim Diff As Double
Dim Denom As Double
Dim tval As Double
Dim k As Integer

    NumLinearità = 9

    With myCertificate.LotCalculation
        Maxx = .MaxValue
        Minx = .MinValue
        
        a = .a
        b = .b
        tval = .tval
    End With
    
    Grid5.Rows = 1
    Grid5.AutoRedraw = False
      
    With myCertificate.GraphCert
      
      .Replicat = 1
      
      ReDim .x(NumLinearità)
      ReDim .Ypred(NumLinearità)
      ReDim .Lplim(NumLinearità)
      ReDim .sYpre(NumLinearità)
      ReDim .Uplim(NumLinearità)
      
      For i = 1 To NumLinearità
        '=2*(MAX(X)-MIN(X))/8+MIN(X)
        .x(i) = (i - 1) * (Maxx - Minx) / (NumLinearità - 1) + Minx
        .Ypred(i) = a + b * .x(i)
        
        'sy*(1/q+1/n+((P6-AVERAGE(X))^2/SSx))^0,5
        Diff = .x(i) - myCertificate.LotCalculation.MedValue
        Denom = (myCertificate.LotCalculation.ssx)
        
        .sYpre(i) = (myCertificate.LotCalculation.sy) * Sqr((1 / .Replicat) + (1 / myCertificate.LotCalculation.n) + Diff ^ 2 / Denom)
          
        'Q6 -tval * R6
        .Lplim(i) = .Ypred(i) - tval * .sYpre(i)
        .Uplim(i) = .Ypred(i) + tval * .sYpre(i)
      

        Grid5.AddItem "", False
        Grid5.cell(i, 0).text = i
        Grid5.cell(i, 1).text = FormatNumber(.x(i), 3)
        Grid5.cell(i, 2).text = FormatNumber(.Ypred(i), 3)
        Grid5.cell(i, 3).text = FormatNumber(.sYpre(i), 3)
        Grid5.cell(i, 4).text = FormatNumber(.Lplim(i), 3)
        Grid5.cell(i, 5).text = FormatNumber(.Uplim(i), 3)
            
        Grid5.cell(i, 4).BackColor = vbColorAzzurrino
        Grid5.cell(i, 5).BackColor = vbColorAzzurrino
      
      Next
      
     ' per il grafico devono essere tanti quanti gli STD.
    Dim Count As Integer
    Dim Ypred As Double
    Dim sYpre As Double
    Count = 0
    
    ReDim .LplimGrph(0)
    ReDim .UplimGrph(0)
    
    For k = 1 To STDCount
        If STD(k, 1) = 0 Then
            GoTo cont:
        End If
        Count = Count + 1
        Ypred = a + b * STD(k, 1)
        
        Diff = STD(k, 1) - myCertificate.LotCalculation.MedValue
        Denom = (myCertificate.LotCalculation.ssx)
        sYpre = (myCertificate.LotCalculation.sy) * Sqr((1 / .Replicat) + (1 / myCertificate.LotCalculation.n) + Diff ^ 2 / Denom)
          
        ReDim Preserve .LplimGrph(Count)
        ReDim Preserve .UplimGrph(Count)
        .LplimGrph(Count) = Ypred - tval * sYpre
        .UplimGrph(Count) = Ypred + tval * sYpre

      
cont:
    Next
    
      Grid5.Column(0).Width = 50
      
      Grid5.ReadOnly = True
      Grid5.Refresh
      Grid5.AutoRedraw = True

    End With

    

    LoadGraph



End Function








Private Function SaveMe()


    Call SetCertificateDescriptionOnFile
    Call SetLotResultOnFile
    Call SetCalibrationFunctionOnFile
    Call SetLotCalculationOnFile
    
    
    PopupMessage 2, "Lot Certificate Saved...", , , "LOT : " & myCertificate.LotNumber

    frCommandInside(1).Visible = True
    

End Function



Private Function SetLotCalculationOnFile()

Dim i As Integer


Dim strA As String
Dim strB As String


    With myCertificate.GraphCert

        
        CloseSettingDataFile
        
            SaveSettingData SettingName, "Certificate", "GrphCount", UBound(.LplimGrph)
            
        For i = 1 To UBound(.LplimGrph)
            
            strA = CStr(FormatNumber(.LplimGrph(i), 3))
            strB = CStr(FormatNumber(.UplimGrph(i), 3))
            SaveSettingData SettingName, "Certificate", "LplimGrph" & i, strA
            SaveSettingData SettingName, "Certificate", "UplimGrph" & i, strB
            SaveSettingData SettingName, "Certificate", "TargetValue" & i, myCertificate.STD(i).STDValue
       Next
        
    
        CloseSettingDataFile
       
    End With


End Function


Private Function SetCertificateDescriptionOnFile()

    With myCertificate
        .ProductName = Grd2.cell(1, 2).text
        .ProductCode = Grd2.cell(2, 2).text
        .Method = Grd2.cell(3, 2).text
        .RangePPM = Grd2.cell(4, 2).text
        .LotNumber = Grd2.cell(5, 2).text
        .BestUseBefore = Grd2.cell(6, 2).text
        .DateAnalisys = Grd2.cell(7, 2).text
        .ReferenceMeter = Grd2.cell(8, 2).text
        .ReferenceSTD = Grd2.cell(9, 2).text
        .Wavelenght = Grd2.cell(10, 2).text
        .CellMM = Grd2.cell(11, 2).text
        .RefSTDNote1 = Grd2.cell(12, 2).text
        .RefSTDNote2 = Grd2.cell(13, 2).text
        .ReferenceMeterDescription = Grd2.cell(14, 2).text
        .RangeFormula = Grd2.cell(15, 2).text
        
        
        CloseSettingDataFile
        
        SaveSettingData SettingName, "Certificate", "TimeStamp", Now()
        SaveSettingData SettingName, "Certificate", "Operator", MyOperatore.Name
        SaveSettingData SettingName, "Certificate", "Soft.Release", App.Major & "." & App.Minor & "." & App.Revision
        
        SaveSettingData SettingName, "Certificate", "ProductName", .ProductName
        SaveSettingData SettingName, "Certificate", "ProductCode", .ProductCode
        SaveSettingData SettingName, "Certificate", "Method", .Method
        SaveSettingData SettingName, "Certificate", "RangePPM", .RangePPM
        SaveSettingData SettingName, "Certificate", "LotNumber", .LotNumber
        SaveSettingData SettingName, "Certificate", "BestUseBefore", .BestUseBefore
        SaveSettingData SettingName, "Certificate", "DateAnalisys", .DateAnalisys
        SaveSettingData SettingName, "Certificate", "ReferenceMeter", .ReferenceMeter
        SaveSettingData SettingName, "Certificate", "ReferenceSTD", .ReferenceSTD
        SaveSettingData SettingName, "Certificate", "Wavelenght", .Wavelenght
        SaveSettingData SettingName, "Certificate", "CellMM", .CellMM
        SaveSettingData SettingName, "Certificate", "RefSTDNote1", .RefSTDNote1
        SaveSettingData SettingName, "Certificate", "RefSTDNote2", .RefSTDNote2
        SaveSettingData SettingName, "Certificate", "ReferenceMeterDescription", .ReferenceMeterDescription
        SaveSettingData SettingName, "Certificate", "UserDecimal", .UserDecimal
        SaveSettingData SettingName, "Certificate", "RangeFormula", .RangeFormula
        
        CloseSettingDataFile
       
    End With


End Function

Private Function SetLotResultOnFile()
Dim k As Integer

   CloseSettingDataFile
   
    With Grid1
    
        SaveSettingData SettingName, "Certificate - Lot Result", "Rows", .Rows - 1
        
        For k = 1 To .Rows - 1
    
             SaveSettingData SettingName, "Certificate - Lot Result", "StdValue" & k, .cell(k, 1).text
             SaveSettingData SettingName, "Certificate - Lot Result", "AverageResult" & k, .cell(k, 2).text
        Next
     
    End With
    
    
    With Grid6
    
        For k = 1 To .Cols - 1

             SaveSettingData SettingName, "Certificate - Components identification", "Code #" & k, .cell(1, k).text
             SaveSettingData SettingName, "Certificate - Components identification", "Lot #" & k, .cell(2, k).text
             SaveSettingData SettingName, "Certificate - Components identification", "Exp #" & k, .cell(3, k).text
        Next
     
    End With
    
     
    

    CloseSettingDataFile
       

End Function

Private Function SetCalibrationFunctionOnFile()
Dim k As Integer
Dim i As Integer

   CloseSettingDataFile
   
    With Grid2
    
        SaveSettingData SettingName, "Certificate - Calibration Function", "Rows", .Rows - 1
        SaveSettingData SettingName, "Certificate - Calibration Function", "Cols", .Cols - 1
        
        For k = 0 To .Rows - 1
            For i = 0 To .Cols - 1
                SaveSettingData SettingName, "Certificate - Calibration Function", "Cell(" & k & "," & i & ")", .cell(k, i).text
            Next
        Next
     
    End With

    CloseSettingDataFile
       

End Function










Private Function GetCertificateDescriptionFromFile()


Dim strName As String

    With myCertificate

        CloseSettingDataFile
        
        strName = GetSettingData(SettingName, "Certificate", "ProductName", "")
        
        If strName = "" Then Exit Function
        
        
        
        
        .ProductName = GetSettingData(SettingName, "Certificate", "ProductName", .ProductName)
        .ProductCode = GetSettingData(SettingName, "Certificate", "ProductCode", .ProductCode)
        .Method = GetSettingData(SettingName, "Certificate", "Method", .Method)
        .RangePPM = GetSettingData(SettingName, "Certificate", "RangePPM", .RangePPM)
        .LotNumber = GetSettingData(SettingName, "Certificate", "LotNumber", .LotNumber)
        .BestUseBefore = GetSettingData(SettingName, "Certificate", "BestUseBefore", .BestUseBefore)
        .DateAnalisys = GetSettingData(SettingName, "Certificate", "DateAnalisys", .DateAnalisys)
        .ReferenceMeter = GetSettingData(SettingName, "Certificate", "ReferenceMeter", .ReferenceMeter)
        .ReferenceMeterDescription = GetSettingData(SettingName, "Certificate", "ReferenceMeterDescription", .ReferenceMeterDescription)
        .ReferenceSTD = GetSettingData(SettingName, "Certificate", "ReferenceSTD", .ReferenceSTD)
        .Wavelenght = GetSettingData(SettingName, "Certificate", "Wavelenght", .Wavelenght)
        .CellMM = GetSettingData(SettingName, "Certificate", "CellMM", .CellMM)
        .RefSTDNote1 = GetSettingData(SettingName, "Certificate", "RefSTDNote1", .RefSTDNote1)
        .RefSTDNote2 = GetSettingData(SettingName, "Certificate", "RefSTDNote2", .RefSTDNote2)
        .UserDecimal = GetSettingData(SettingName, "Certificate", "UserDecimal", .UserDecimal)
        .RangeFormula = GetSettingData(SettingName, "Certificate", "RangeFormula", .RangeFormula)
        
    
        CloseSettingDataFile
        
        If .ProductName <> "" And .ProductCode <> "" Then
        
        
            Grd2.cell(1, 2).text = .ProductName
            Grd2.cell(2, 2).text = .ProductCode
            Grd2.cell(3, 2).text = .Method
            Grd2.cell(4, 2).text = .RangePPM
            Grd2.cell(5, 2).text = .LotNumber
            Grd2.cell(6, 2).text = .BestUseBefore
            Grd2.cell(7, 2).text = .DateAnalisys
            Grd2.cell(8, 2).text = .ReferenceMeter
            Grd2.cell(9, 2).text = .ReferenceSTD
            Grd2.cell(10, 2).text = .Wavelenght
            Grd2.cell(11, 2).text = .CellMM
            Grd2.cell(12, 2).text = .RefSTDNote1
            Grd2.cell(13, 2).text = .RefSTDNote2
            Grd2.cell(14, 2).text = .ReferenceMeterDescription
            Grd2.cell(15, 2).text = .RangeFormula
        
        End If
        
        
        Label1(0) = "Lot Result [" & .RangeFormula & "]"
                    
        If Grid2.Rows > 1 Then
            
            Grid2.cell(2, 0).text = "Standard Deviation [" & .RangeFormula & "]"
            Grid2.cell(4, 0).text = "Confidence interval (95%)[" & .RangeFormula & "]"
            Grid2.AutoFitRowHeight (2)
            Grid2.AutoFitRowHeight (4)
        End If

        
       
    End With


End Function

Private Function ReagentSet(ByVal rc As Boolean)
Dim i As Integer
Dim NumCount As Integer

CloseSettingDataFile

NumCount = IIf(rc, 10, 44)

With Grid6

    ReDim iReagentSet(5)
    .DefaultRowHeight = 20
    .DefaultFont.Size = 8
    .cell(1, 0).text = "Component code "
    .cell(2, 0).text = "Lot Number "
    .cell(3, 0).text = "Expiration "
    
    .Column(0).AutoFit
    .Column(0).Alignment = cellLeftCenter
    
    For i = 1 To 5
        With iReagentSet(i)
             .Code = GetSettingData(SettingName, "Information QC", "Reagent Code" & NumCount + i, "")
             .Lot = GetSettingData(SettingName, "Information QC", "Text1" & NumCount + i, "")
             .Exp = GetSettingData(SettingName, "Information QC", "Text1" & NumCount + i + 5, "")
             
             If .Lot = "" Or .Code = "" Then
                .Code = ""
                .Lot = ""
                .Exp = ""
             End If
             
             
      

             .Code = GetSettingData(SettingName, "Certificate - Components identification", "Code #" & i, .Code)
             .Lot = GetSettingData(SettingName, "Certificate - Components identification", "Lot #" & i, .Lot)
             .Exp = GetSettingData(SettingName, "Certificate - Components identification", "Exp #" & i, .Exp)
      
        End With
        
  
    
        
        .cell(1, i).text = iReagentSet(i).Code
        .cell(2, i).text = iReagentSet(i).Lot
        .cell(3, i).text = iReagentSet(i).Exp
       
        .Column(i).AutoFit
        
    Next
   
 End With
CloseSettingDataFile


End Function


Private Function GetMaxStdDeviation()
Dim i As Integer
Dim k As Integer
Dim MaxStdDeviation As Double
Dim TempStDev As Double
CloseSettingDataFile
MaxStdDeviation = 0
For k = 1 To STDCount
    TempStDev = GetSettingData(SettingName, "Evaluation QC", "StdDeviation" & k, 0)
    MaxStdDeviation = IIf(MaxStdDeviation < TempStDev, TempStDev, MaxStdDeviation)
    'SaveSettingData SettingName, "Evaluation QC", "StdDeviatioPerc" & STDNumber, StdDeviatioPerc
   ' SaveSettingData SettingName, "Evaluation QC", "Repeatability" & STDNumber, Repeatability
Next


myCertificate.CalibrationFunction.StdDeviation.LotValue = MaxStdDeviation
CloseSettingDataFile

End Function


Private Sub TimerIntro_Timer()

TimerIntro.Enabled = False
LotCalculations

End Sub
