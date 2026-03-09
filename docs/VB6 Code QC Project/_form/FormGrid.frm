VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form FormGrid 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Database"
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
   Icon            =   "FormGrid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   12000
   ScaleWidth      =   19200
   Begin VB.Frame frExcel 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3255
      Left            =   1200
      TabIndex        =   58
      Top             =   3960
      Visible         =   0   'False
      Width           =   12015
      Begin VB.Label labell 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Export may take several minutes"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   0
         Left            =   0
         MouseIcon       =   "FormGrid.frx":33E2
         MousePointer    =   99  'Custom
         TabIndex        =   60
         Top             =   2040
         Width           =   11925
      End
      Begin VB.Image image 
         Height          =   480
         Index           =   7
         Left            =   5760
         MouseIcon       =   "FormGrid.frx":36EC
         MousePointer    =   99  'Custom
         Picture         =   "FormGrid.frx":39F6
         Top             =   720
         Width           =   480
      End
      Begin VB.Label labell 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please do not close Chemical QC during Excel Operations...."
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
         Left            =   0
         MouseIcon       =   "FormGrid.frx":6DD8
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   1560
         Width           =   11940
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5775
      Left            =   2880
      TabIndex        =   32
      Top             =   1800
      Visible         =   0   'False
      Width           =   13695
      Begin FlexCell.Grid Grd3 
         Height          =   3600
         Left            =   480
         TabIndex        =   33
         Top             =   1320
         Width           =   12720
         _ExtentX        =   22437
         _ExtentY        =   6350
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
         BorderColor     =   -2147483635
         CellBorderColor =   16512
         CellBorderColorFixed=   16777215
         Cols            =   10
         DefaultFontName =   "Calibri"
         DefaultFontSize =   12
         DisplayDateTimeMask=   -1  'True
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   16777215
         ReadOnly        =   -1  'True
         Rows            =   10
         SelectionMode   =   1
         MultiSelect     =   0   'False
         DateFormat      =   2
         EnterKeyMoveTo  =   1
         BackColorComment=   -2147483635
         AllowUserPaste  =   2
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   6600
         Picture         =   "FormGrid.frx":70E2
         Top             =   5060
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Specifications Lot :"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2520
         MouseIcon       =   "FormGrid.frx":A4C4
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   720
         Width           =   1920
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   2
         Left            =   1920
         Picture         =   "FormGrid.frx":A7CE
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lbSpecification 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOT_NUMBER"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   4560
         MouseIcon       =   "FormGrid.frx":DBB0
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   585
         Width           =   7215
      End
   End
   Begin VB.PictureBox PicMainMenu 
      BackColor       =   &H00303030&
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
      TabIndex        =   21
      Top             =   0
      Width           =   19215
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   6
         Left            =   11520
         MouseIcon       =   "FormGrid.frx":DEBA
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   6
            Left            =   720
            MouseIcon       =   "FormGrid.frx":E1C4
            MousePointer    =   99  'Custom
            Picture         =   "FormGrid.frx":E4CE
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Close Lot"
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
            Index           =   6
            Left            =   0
            MouseIcon       =   "FormGrid.frx":10EC0
            MousePointer    =   99  'Custom
            TabIndex        =   57
            Top             =   720
            Width           =   1890
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   5
         Left            =   5760
         MouseIcon       =   "FormGrid.frx":111CA
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   49
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Export Lot Excel"
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
            Index           =   5
            Left            =   0
            MouseIcon       =   "FormGrid.frx":114D4
            MousePointer    =   99  'Custom
            TabIndex        =   50
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   5
            Left            =   720
            MouseIcon       =   "FormGrid.frx":117DE
            MousePointer    =   99  'Custom
            Picture         =   "FormGrid.frx":11AE8
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   4
         Left            =   9600
         MouseIcon       =   "FormGrid.frx":14ECA
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Mean Value"
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
            Left            =   0
            MouseIcon       =   "FormGrid.frx":151D4
            MousePointer    =   99  'Custom
            TabIndex        =   38
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   4
            Left            =   720
            MouseIcon       =   "FormGrid.frx":154DE
            MousePointer    =   99  'Custom
            Picture         =   "FormGrid.frx":157E8
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   0
         MouseIcon       =   "FormGrid.frx":18BCA
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   28
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "FormGrid.frx":18ED4
            MousePointer    =   99  'Custom
            Picture         =   "FormGrid.frx":191DE
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Refresh List"
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
            Left            =   525
            MouseIcon       =   "FormGrid.frx":1C5C0
            MousePointer    =   99  'Custom
            TabIndex        =   29
            Top             =   720
            Width           =   960
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   1920
         MouseIcon       =   "FormGrid.frx":1C8CA
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   26
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   720
            MouseIcon       =   "FormGrid.frx":1CBD4
            MousePointer    =   99  'Custom
            Picture         =   "FormGrid.frx":1CEDE
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Record Book Excel"
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
            Left            =   210
            MouseIcon       =   "FormGrid.frx":202C0
            MousePointer    =   99  'Custom
            TabIndex        =   27
            Top             =   720
            Width           =   1470
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   3840
         MouseIcon       =   "FormGrid.frx":205CA
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   24
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   720
            MouseIcon       =   "FormGrid.frx":208D4
            MousePointer    =   99  'Custom
            Picture         =   "FormGrid.frx":20BDE
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Date filter"
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
            MouseIcon       =   "FormGrid.frx":23FC0
            MousePointer    =   99  'Custom
            TabIndex        =   25
            Top             =   720
            Width           =   1890
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   3
         Left            =   7680
         MouseIcon       =   "FormGrid.frx":242CA
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   22
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   3
            Left            =   720
            MouseIcon       =   "FormGrid.frx":245D4
            MousePointer    =   99  'Custom
            Picture         =   "FormGrid.frx":248DE
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Delete Lot"
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
            MouseIcon       =   "FormGrid.frx":27CC0
            MousePointer    =   99  'Custom
            TabIndex        =   23
            Top             =   720
            Width           =   1890
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Lot List"
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
         Left            =   15525
         TabIndex        =   30
         Top             =   360
         Width           =   3165
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   2880
      ScaleHeight     =   2895
      ScaleWidth      =   13455
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   13455
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel filter"
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
         Left            =   12210
         MouseIcon       =   "FormGrid.frx":27FCA
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   2580
         Width           =   990
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apply filter"
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
         Left            =   240
         MouseIcon       =   "FormGrid.frx":282D4
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   2580
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Date Range and Apply Filter"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   4560
         TabIndex        =   31
         Top             =   2160
         Width           =   4500
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   11
         Left            =   480
         MouseIcon       =   "FormGrid.frx":285DE
         MousePointer    =   99  'Custom
         Picture         =   "FormGrid.frx":288E8
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   10
         Left            =   12480
         MouseIcon       =   "FormGrid.frx":2BCCA
         MousePointer    =   99  'Custom
         Picture         =   "FormGrid.frx":2BFD4
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lbDataFiltro 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   855
         Index           =   1
         Left            =   7560
         MouseIcon       =   "FormGrid.frx":2F3B6
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label lbDataFiltro 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   855
         Index           =   0
         Left            =   1560
         MouseIcon       =   "FormGrid.frx":2F6C0
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "END DATE"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         Left            =   8970
         TabIndex        =   13
         Top             =   480
         Width           =   1440
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "START DATE"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   2820
         TabIndex        =   12
         Top             =   480
         Width           =   1785
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   11
         Left            =   0
         TabIndex        =   18
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1095
         Index           =   10
         Left            =   11760
         TabIndex        =   17
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00303030&
         Height          =   2895
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   13455
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   2040
      Top             =   1800
   End
   Begin VB.Timer Timer3 
      Interval        =   250
      Left            =   1080
      Top             =   1920
   End
   Begin VB.ComboBox CmbVisual 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   510
      Left            =   14640
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   10020
      Width           =   4095
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   510
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   10020
      Width           =   4095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00008000&
      Caption         =   "Frame3"
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   17400
      TabIndex        =   2
      Top             =   6480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   1560
      Top             =   1920
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   10680
      MouseIcon       =   "FormGrid.frx":2F9CA
      ScaleHeight     =   855
      ScaleWidth      =   4455
      TabIndex        =   6
      Top             =   10880
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   8
         Left            =   2040
         MouseIcon       =   "FormGrid.frx":2FCD4
         MousePointer    =   99  'Custom
         Picture         =   "FormGrid.frx":2FFDE
         Top             =   200
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin ChemicalQC.ctlCalendar ctlCalendar1 
      Height          =   6960
      Left            =   2520
      TabIndex        =   20
      Top             =   4320
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
   Begin FlexCell.Grid GrdLot 
      Height          =   7680
      Left            =   240
      TabIndex        =   36
      Top             =   1800
      Width           =   18720
      _ExtentX        =   33020
      _ExtentY        =   13547
      AllowUserReorderColumn=   -1  'True
      AllowUserSort   =   -1  'True
      Appearance      =   0
      BackColor1      =   15790320
      BackColor2      =   15790320
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
      DefaultFontSize =   9.75
      BoldFixedCell   =   0   'False
      ButtonLocked    =   -1  'True
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
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   2
      Left            =   0
      MouseIcon       =   "FormGrid.frx":333C0
      MousePointer    =   99  'Custom
      TabIndex        =   48
      Top             =   10680
      Width           =   2055
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   1
      Left            =   3720
      MouseIcon       =   "FormGrid.frx":336CA
      MousePointer    =   99  'Custom
      TabIndex        =   47
      Top             =   10680
      Width           =   2055
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   0
      Left            =   8640
      MouseIcon       =   "FormGrid.frx":339D4
      MousePointer    =   99  'Custom
      TabIndex        =   46
      Top             =   10680
      Width           =   2055
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   3
      Left            =   15000
      MouseIcon       =   "FormGrid.frx":33CDE
      MousePointer    =   99  'Custom
      TabIndex        =   45
      Top             =   10680
      Width           =   2055
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   4
      Left            =   17400
      MouseIcon       =   "FormGrid.frx":33FE8
      MousePointer    =   99  'Custom
      TabIndex        =   44
      Top             =   10680
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Lots"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   2
      Left            =   3870
      MouseIcon       =   "FormGrid.frx":342F2
      MousePointer    =   99  'Custom
      TabIndex        =   55
      Top             =   1320
      Width           =   870
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Closed Lots"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   1920
      MouseIcon       =   "FormGrid.frx":345FC
      MousePointer    =   99  'Custom
      TabIndex        =   54
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open Lots"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   270
      MouseIcon       =   "FormGrid.frx":34906
      MousePointer    =   99  'Custom
      TabIndex        =   53
      Top             =   1320
      Width           =   1140
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   12
      Left            =   17760
      MouseIcon       =   "FormGrid.frx":34C10
      MousePointer    =   99  'Custom
      TabIndex        =   43
      Top             =   11600
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   11
      Left            =   15360
      MouseIcon       =   "FormGrid.frx":34F1A
      MousePointer    =   99  'Custom
      TabIndex        =   42
      Top             =   11600
      Width           =   1230
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Database"
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
      Index           =   7
      Left            =   9030
      MouseIcon       =   "FormGrid.frx":35224
      MousePointer    =   99  'Custom
      TabIndex        =   41
      Top             =   11600
      Width           =   1110
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apply filter"
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
      Index           =   6
      Left            =   4335
      MouseIcon       =   "FormGrid.frx":3552E
      MousePointer    =   99  'Custom
      TabIndex        =   40
      Top             =   11600
      Width           =   930
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Excel folder"
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
      Index           =   5
      Left            =   450
      MouseIcon       =   "FormGrid.frx":35838
      MousePointer    =   99  'Custom
      TabIndex        =   39
      Top             =   11600
      Width           =   960
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Hi340504333"
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
      Height          =   495
      Index           =   0
      Left            =   4800
      MouseIcon       =   "FormGrid.frx":35B42
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   10035
      Width           =   9585
   End
   Begin VB.Label lbColonneGrid 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show Less"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   17760
      MouseIcon       =   "FormGrid.frx":35E4C
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   1320
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIOD"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   1
      Left            =   14640
      TabIndex        =   10
      Top             =   9600
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODE"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   0
      Left            =   4815
      TabIndex        =   9
      Top             =   9600
      Width           =   9570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FILTER"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   5
      Left            =   480
      TabIndex        =   8
      Top             =   9600
      Width           =   915
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FormGrid.frx":36156
      Top             =   9360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   4
      Left            =   18240
      MouseIcon       =   "FormGrid.frx":39538
      MousePointer    =   99  'Custom
      Picture         =   "FormGrid.frx":39842
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   15600
      MouseIcon       =   "FormGrid.frx":3CC24
      MousePointer    =   99  'Custom
      Picture         =   "FormGrid.frx":3CF2E
      Top             =   11040
      Width           =   480
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
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   600
      MouseIcon       =   "FormGrid.frx":40310
      MousePointer    =   99  'Custom
      Picture         =   "FormGrid.frx":4061A
      Top             =   11040
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   480
      X2              =   18720
      Y1              =   10680
      Y2              =   10680
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   1
      Left            =   4560
      MouseIcon       =   "FormGrid.frx":439FC
      MousePointer    =   99  'Custom
      Picture         =   "FormGrid.frx":43D06
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   0
      Left            =   9360
      MouseIcon       =   "FormGrid.frx":470E8
      MousePointer    =   99  'Custom
      Picture         =   "FormGrid.frx":473F2
      Top             =   11040
      Width           =   480
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1575
      Index           =   6
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Index           =   5
      Left            =   17280
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   8
      Left            =   11640
      TabIndex        =   5
      Top             =   10560
      Width           =   2655
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1095
      Index           =   9
      Left            =   3960
      TabIndex        =   16
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "FormGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private bModMenu As Boolean
Private bCalcoloTitolo As Boolean
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
Private ExcelFilename As String

Private IndexTabella As Integer
Private MaxIndex As Integer
Private dIndexProcedura As Integer
Private m_Procedura As Boolean
Private CampioneSelezionato As String

Private bHilight As Boolean
Private bFiltroIntervalloDate As Boolean
Private DataIndex As Integer
Private m_rc As Boolean

Private MyDA As String
Private MyA As String
Private MyPeriodo As String
Private MyID As Long
Private MyFileName As String
Private bMeanValue As Boolean
Private bClosedLot As Boolean
Private MyIndexRecord As Integer
Private lRow As Long
Private IndexOpenClosedLot As Integer
Private sCode As String
Private bScheduleExcel As Boolean
Private UNIT_PP As String
Private MeasurementUnit As String


Private Sub Form_Load()
If bFullScreen Then Me.WindowState = 2


IndexOpenClosedLot = 2
ChangeLabelLots

Call SetGrid(GrdLot)

Call GrdRisultati(Grd3, "ppm")

    MyID = 0
    MyFileName = ""
    MyIndexRecord = 3
    bMeanValue = False
    bClosedLot = False
    
    RiempiCombo
    
    SaveSizes
End Sub


Public Function DoShow(Optional ByVal IndexProcedura As Integer = 0, Optional ByRef MyCampione As String = "") As Boolean
Dim i As Integer
    On Error GoTo ERR_SHOW
    
    m_rc = False
    mOk
    dIndexProcedura = IndexProcedura
    
    CampioneSelezionato = MyCampione
    

    

    Me.Show vbModal
    
    

    
    If m_rc = True Then
        MyCampione = CampioneSelezionato
    End If
ERR_END:
    On Error GoTo 0
    DoShow = m_rc
    Exit Function
ERR_SHOW:
    m_rc = False
    Resume ERR_END
End Function




Private Sub Frame1_Click()
Frame1.Visible = False
End Sub

Private Sub Grd3_dblClick()
Frame1.Visible = False
End Sub

Private Sub GrdLot_BeforeUserSort(ByVal Col As Long)
lRow = 0
End Sub

Private Sub GrdLot_Click()
ctlCalendar1.Visible = False

End Sub

Private Sub Image2_Click()
Frame1.Visible = False
End Sub

Private Sub Image3_Click(Index As Integer)

Dim MyexcelName As String
    Select Case Index
        Case 0
            Combo1.ListIndex = 0
            Combo1_Click
             IntervalloDate False
            GlobalSearch
        Case 1
              ' stampa lista
        
            'USER_EXCEL_PATH
            If F_InputBox.DoShow("Please Enter Excel File Name", "RecordBook", , , , MyexcelName) Then
                If MyexcelName = "" Then
                    MyexcelName = USER_EXCEL_PATH & "\" & FormatNomeFile("RecordBook" & Now & ".xls")
                Else
                    MyexcelName = USER_EXCEL_PATH & "\" & FormatNomeFile(MyexcelName & ".xls")
                End If
                
            Else
                Exit Sub
            End If
            
            GrdLot.ExportToExcel MyexcelName, True, True
            MessageInfoTime = 2000
            PopupMessage 2, ("PDF Report saved in : ") & vbCrLf & MyexcelName, , , "Record Book"

        Case 2
            Call OpenIntervalloDate
        Case 3
             ' cancella campione
             If CheckPrivilege(3) Then
                If InStr(Text1(0), Combo1) Then
                    GoTo err_canc:
                Else
                    If Text1(0) = "" Then
err_canc:
                        If MyID > 0 Then Call CancellaTab
                        
                    Else
                        Call CancellaTab
                    End If
                End If
            End If
        
        Case 4
            ' mean value
            If lRow > 0 Then
                If bMeanValue Then
                
                    Frame1.Visible = Not (Frame1.Visible)
                End If
            End If
            
        Case 5
            ' export LOT Excel
            Call ExportAllExcel
            
           
            
        Case 6
        ' cancella campione
             If CheckPrivilege(3) Then
                If InStr(Text1(0), Combo1) Then
                   
                Else
                    If Text1(0) = "" Then

                    Else
                        Call ChiudiTab
                    End If
                End If
            End If
    End Select
End Sub
Private Function ExportAllExcel()


frExcel.Left = Me.Width / 2 - frExcel.Width / 2
frExcel.Top = Me.Height / 2 - frExcel.Height / 2
If bScheduleExcel = False Then

If F_MsgBox.DoShow("Do you want to schedule more than one file?", "Export to Excel") Then

    bScheduleExcel = True
    PopupMessage 2, "Select rows and press Excel", , , "Schedule Excel Reports"
Else

    frExcel.Visible = True

    PopupMessage 2, "Exporting " & lbSpecification.Caption & " : please wait...."

    
    ExportExcelOneFile

    frExcel.Visible = False
            
End If
            
            
Else
    
    If F_MsgBox.DoShow("Export all records to Excel?", "Export to Excel") Then
    
        frExcel.Visible = True
        With GrdLot
            Dim i As Integer
            Dim t As Integer
            For i = 1 To .Rows - 1
                If .Cell(i, 1).BackColor = vbColorAzzurrino Then
                               
                        ' esporta in excel.....
                        MyFileName = .Cell(i, 16).Text
                        ExcelFilename = Trim(.Cell(i, 4).Text) & "_" & Trim(.Cell(i, 2).Text) & "_" & Trim(.Cell(i, 6).Text) & "_LOT" & Trim(.Cell(i, 1).Text & "_PW" & Trim(.Cell(i, 5).Text))
    
   
                        If ExportExcelOneFile Then
                        
                            For t = 1 To .Cols - 3
                                
                                  
                                .Cell(i, t).BackColor = &HF0F0F0
                                
                            
                            Next
                        
                        
                        End If
                        
                        
             
                End If
            Next
            
            
            
            
        End With
        Image3_Click 0
        frExcel.Visible = False
        bScheduleExcel = False
    Else
        
        For i = 1 To GrdLot.Rows - 1
              For t = 1 To GrdLot.Cols - 3
                                
                      
                    GrdLot.Cell(i, t).BackColor = &HF0F0F0
                    
                
                Next
        
        
        Next
    
        bScheduleExcel = False
    End If



End If


End Function


Private Function ExportExcelOneFile() As Boolean
Dim rc As Boolean
Dim MyexcelName As String

rc = True

   
    If MyFileName = "" Then
        rc = False
    Else
        USER_PATH = IIf(bClosedLot, USER_DATA_PATH, USER_TEMP_PATH)
        
        If Len(ExcelFilename) > 50 Then ExcelFilename = Left$(ExcelFilename, 49)
        
        MyexcelName = FormatNomeFile(ExcelFilename)
        
      
        Call EsportaExcel(MyFileName, MyexcelName)
       ' USER_PATH = USER_TEMP_PATH
    End If
    

    ExportExcelOneFile = rc
    

End Function
Private Sub Label2_Click(Index As Integer)
Image3_Click Index
End Sub

Private Sub Label4_Click(Index As Integer)
IndexOpenClosedLot = Index
ChangeLabelLots
GlobalSearch
End Sub

Private Sub lbColonneGrid_Click()
Dim rc As Boolean
rc = IIf(GrdLot.Column(3).Width = 0, True, False)
SaveSetting App.Title, "Settings Filtro", "Visualizza Colonne", Not (rc)
VisulaizzaColonne rc

End Sub

Private Sub VisulaizzaColonne(ByVal rc As Boolean)
With GrdLot


          ' .Cell(0, 0).Text = "n."
       ' .Cell(0, 1).Text = "Lot Number"
       ' .Cell(0, 2).Text = "Code SFG"
       ' .Cell(0, 3).Text = "Description"
       ' .Cell(0, 4).Text = "Recipe"
       ' .Cell(0, 5).Text = "Prep. Week"
       ' .Cell(0, 6).Text = "Range Min"
       ' .Cell(0, 7).Text = "Range Max"
       ' .Cell(0, 8).Text = "Date"
       ' .Cell(0, 9).Text = "Exp.Date"
       ' .Cell(0, 10).Text = "# Test" ' quanti test ho fatto
       ' .Cell(0, 11).Text = "Mean Value" ' se ho fatto calcolo medie
       ' .Cell(0, 12).Text = "Finalise" ' se ho finalizzat ( solo Laboratory Manager )
       ' .Cell(0, 13).Text = "QC Operator"
       ' .Cell(0, 14).Text = "QC Note"
       ' .Cell(0, 15).Text = "ID"
       ' .Cell(0, 16).Text = "FileName"
       ' .Cell(0, 17).Text = "NomeFileReport"
       ' .Cell(0, 18).Text = "NomeFileExcel"
       ' .Cell(0, 19).Text = "CODE_ID"
       
    .Column(3).Width = IIf(rc, 350 * m_ControlGridColWidth, 0)
    .Column(6).Width = IIf(rc, 120 * m_ControlGridColWidth, 0)
    .Column(7).Width = IIf(rc, 120 * m_ControlGridColWidth, 0)
    .Column(8).Width = IIf(rc, 120 * m_ControlGridColWidth, 0)
    
    .Column(13).Width = IIf(rc, 350 * m_ControlGridColWidth, 0)
    .Column(14).Width = 0 'IIf(rc, 500 * m_ControlGridColWidth, 0)
    
End With
lbColonneGrid.Caption = IIf(rc, ("Hide Columns"), ("Show Columns"))
End Sub
Private Sub lbDataFiltro_Change(Index As Integer)
    MessageInfoTime = 2000
    Select Case Index
        Case 0
            If Len(lbDataFiltro(1)) > 0 And Len(lbDataFiltro(0)) > 0 Then
                ctlCalendar1.Visible = False
                If CDate(lbDataFiltro(0)) > CDate(lbDataFiltro(1)) Then
                    PopupMessage 2, ("La data di inizio deve essere più recente o uguale alla data di fine."), , True, ("Filter"), DefaultMenu(9)
                    lbDataFiltro(0) = ""
                End If
                
            End If
        Case 1
            If Len(lbDataFiltro(0)) > 0 And Len(lbDataFiltro(1)) > 0 Then
                ctlCalendar1.Visible = False
                If CDate(lbDataFiltro(0)) > CDate(lbDataFiltro(1)) Then
                    
                    PopupMessage 2, ("La data di Fine non può essere più recente della data di inizio."), , True, ("Filter"), DefaultMenu(9)
                    lbDataFiltro(1) = ""
                End If
                                
                
            End If
    
    
    End Select
    If Index = 0 Or Index = 1 Then
        lbDataFiltro(Index).BackColor = IIf(Len(lbDataFiltro(Index)) > 0, vbWhite, &H8000000D)
    End If
End Sub


Private Sub Form_Resize()
ResizeControls
ResizeTab
End Sub


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


For Each Ctl In Controls
    With m_ControlPositions(i)
        If TypeOf Ctl Is Line Then
            Ctl.X1 = x_scale * .Left
            Ctl.Y1 = y_scale * .Top
            Ctl.X2 = Ctl.X1 + x_scale * .Width
            Ctl.Y2 = Ctl.Y1 + y_scale * .Height
        
        ElseIf TypeOf Ctl Is Grid Then
           Ctl.Left = x_scale * .Left
            Ctl.Top = y_scale * .Top
            Ctl.Width = x_scale * .Width
            Ctl.Height = y_scale * .Height

              
        ElseIf TypeOf Ctl Is Menu Then
        ElseIf TypeOf Ctl Is Timer Then
        Else
            Ctl.Left = x_scale * .Left
            'MsgBox (TypeName(ctl))
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
End Sub

Private Sub CmbVisual_Click()
SaveSetting App.Title, "Settings Filtro", "Filtro Data", CmbVisual.ListIndex

MyDA = ""
MyA = ""
If CmbVisual.ListIndex < 0 Then
    MyPeriodo = 1
Else
    MyPeriodo = GetDateCombo(CmbVisual.ListIndex)

End If


If Me.Visible Then GlobalSearch
End Sub


Private Sub Combo1_Click()
'If Me.Visible Then
GrdLot.Cell(0, 0).SetFocus
    
    Text1(0) = (" - Search") & Combo1 & " - "
    Label1(0) = UCase(Combo1)
    'If Me.Visible Then Text1(0).SetFocus
    SaveSetting App.Title, "Settings Filtro", "Filtro Combo", Combo1.ListIndex
'End If
End Sub




Private Sub CmbVisual_GotFocus()
CmbVisual.ForeColor = vbWhite
End Sub

Private Sub Combo1_GotFocus()
Combo1.ForeColor = vbWhite
End Sub

Private Sub ctlCalendar1_DateClicked(inputDate As Date)
lbDataFiltro(DataIndex) = inputDate
ctlCalendar1.Visible = False
End Sub

Private Sub lbDataFiltro_Click(Index As Integer)
If Index = 0 Or Index = 1 Then
    ctlCalendar1.Left = Picture2.Left + lbDataFiltro(Index).Left + (lbDataFiltro(Index).Width / 2 - ctlCalendar1.Width / 2)
    ctlCalendar1.ZOrder
    ctlCalendar1.Visible = True
    DataIndex = Index
Else
    OpenIntervalloDate
End If

End Sub

Private Sub DefaultMenu_Click(Index As Integer)
Select Case Index
    Case 0
        Unload Me
    Case 2
        ' Open Report folder
        OpenWithDefault (USER_EXCEL_PATH)
      
    Case 1
        ' filtro
        
            Call GlobalSearch
    Case 4
        ' avanti di 10
        Call ScorriTabella(True)
    Case 3
        ' indietro di 10
        Call ScorriTabella(False)
    
    
    
    Case 5
   
    Case 6
      
    Case 7
        ' aggiungi campione
        GetFormDatiCampione
        
    Case 8
        ' play
        'PopupMessage 2, "Campione Importato in procedura", , , FormName
        CampioneSelezionato = Text1(0)
        m_Procedura = True
        m_rc = True
        Unload Me
    Case 9
       
    Case 10
    
        IntervalloDate False
      
    Case 11
        If lbDataFiltro(0) = "" Or lbDataFiltro(1) = "" Then
            PopupMessage 2, ("Select Start and End Filter Date"), , , ("Filter"), DefaultMenu(9)
        Else
        
            IntervalloDate True
        End If
    Case 12

        
End Select
End Sub



Private Sub DefaultMenuLabel_Click(Index As Integer)
DefaultMenu_Click Index
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub



Private Sub RiempiCombo()

    With CmbVisual
        .Clear
        .AddItem "Day"
        .AddItem "Month"
        .AddItem "Year"
        .AddItem "Archive"
        .ListIndex = GetSetting(App.Title, "Settings Filtro", "Filtro Data", 0)
    End With

    With Combo1
        .Clear
        .AddItem " " & ("Code")
        .AddItem " " & ("Description")
        .AddItem " " & ("Lot")
        .AddItem " " & ("Recipe")
        Combo1.ListIndex = 0 ' GetSetting(App.Title, "Settings Filtro", "Filtro Combo", 0)
    End With
    
End Sub



Private Function RiempiGrid(ByRef Grd As Grid)
Dim i As Integer
Dim t As Integer

    ' --------------------------------------
    '
    '  filtra TabReport e riempi Tabella
    '
    ' --------------------------------------
    MyDA = lbDataFiltro(0)
    MyA = lbDataFiltro(1)
    Call FillTabellaTutte(Grd, MyPeriodo, Combo1, Text1(0), MyDA, MyA, IndexOpenClosedLot)
    IndexTabella = 1
    MaxIndex = IIf(Int((GrdLot.Rows - 1) / 10) < (GrdLot.Rows - 1) / 10, (Int((GrdLot.Rows - 1) / 10)) + 1, Int((GrdLot.Rows - 1) / 10))
    If MaxIndex = 0 Then MaxIndex = 1
    
End Function







Private Sub ImageTAV_Click(Index As Integer)
Select Case Index
        Case 0
            Unload Me
        
        Case 2
        

End Select
End Sub

Private Sub ResizeTab()
Dim rc As Boolean
    With Grd3
      .AutoRedraw = False
      
        .DefaultRowHeight = 38 ' * m_ControlGridRowHeight
        .Column(0).Width = 0 '55' * m_ControlGridColWidth
        .Column(1).Width = 200 ' * m_ControlGridColWidth
        .Column(2).Width = 220 ' * m_ControlGridColWidth
        .Column(3).Width = 220 ' * m_ControlGridColWidth
    
        .DefaultFont.Size = 12 '* m_ControlGridFontSize
        .DefaultFont.Bold = False
        


     
        
      .AutoRedraw = True
      .Refresh
      .ReadOnly = True
      
    End With
    With GrdLot
    
        .AutoRedraw = False

        .DefaultRowHeight = 28 ' * m_ControlGridRowHeight
        .Column(0).Width = 33 ' * m_ControlGridColWidth
        .Column(1).Width = 170 ' * m_ControlGridColWidth
        .Column(2).Width = 170 ' * m_ControlGridColWidth
        .Column(3).Width = 350 ' * m_ControlGridColWidth
        .Column(4).Width = 120 ' * m_ControlGridColWidth
        .Column(5).Width = 120 ' * m_ControlGridColWidth
        .Column(6).Width = 120 ' * m_ControlGridColWidth
        .Column(7).Width = 120 ' * m_ControlGridColWidth
        .Column(8).Width = 120 ' * m_ControlGridColWidth
        .Column(9).Width = 150 ' * m_ControlGridColWidth
        .Column(10).Width = 150 ' 150' * m_ControlGridColWidth
        .Column(11).Width = 150 ' * m_ControlGridColWidth
        .Column(12).Width = 170 ' * m_ControlGridColWidth
        .Column(13).Width = 250 ' * m_ControlGridColWidth
        .Column(14).Width = 500 ' * m_ControlGridColWidth
        .DefaultFont.Size = 12 '* m_ControlGridFontSize
        .DefaultFont.Bold = False
        .AutoRedraw = True
        .Refresh
        .ReadOnly = True
    
    
    End With
  
  rc = (GetSetting(App.Title, "Settings Filtro", "Visualizza Colonne", True))
  VisulaizzaColonne Not (rc)
End Sub

Private Sub GrdLot_DblClick()
If lRow > 0 Then
    If bMeanValue Then
        Frame1.Visible = True
    End If
End If

End Sub

Private Sub GrdLot_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim NumCol As Integer

MyID = 0
MyFileName = ""
bMeanValue = False
bClosedLot = False
lRow = FirstRow
lbSpecification = ""
sCode = ""
ExcelFilename = ""
PicMenu(4).Visible = False
PicMenu(6).Visible = False
lRow = 0

If FirstRow > 0 Then
    
    lRow = FirstRow
    
    

    MyLbHelpCount = 0
    lbSpecification.Caption = Trim(GrdLot.Cell(lRow, 1).Text) & " - " & Trim(GrdLot.Cell(lRow, 2).Text)
    
    ' RECIPE + CODE + RANGE + LOT + PREPWK
    ExcelFilename = Trim(GrdLot.Cell(lRow, 4).Text) & "_" & Trim(GrdLot.Cell(lRow, 2).Text) & "_" & Trim(GrdLot.Cell(lRow, 6).Text) & "_LOT" & Trim(GrdLot.Cell(lRow, 1).Text & "_PW" & Trim(GrdLot.Cell(lRow, 5).Text))
    
   
    
    
    NumCol = SetNumCol(Combo1)
    Text1(0) = Trim(GrdLot.Cell(FirstRow, NumCol).Text)
    MyID = GrdLot.Cell(FirstRow, 15).Text
    MyFileName = GrdLot.Cell(FirstRow, 16).Text
    bMeanValue = GrdLot.Cell(FirstRow, 11).Text
    bClosedLot = GrdLot.Cell(FirstRow, 12).Text
    sCode = GrdLot.Cell(FirstRow, 2).Text
    
    
    
     
    If bScheduleExcel Then
    
    
        Dim t As Integer
        Dim UserColor As OLE_COLOR
        
        UserColor = vbColorAzzurrino
        If GrdLot.Cell(lRow, 1).BackColor = UserColor Then
            UserColor = &HF0F0F0
        Else
        
        End If
        
        
        For t = 1 To GrdLot.Cols - 3
          
            GrdLot.Cell(lRow, t).BackColor = UserColor
        Next
        Exit Sub
    End If
    
    
    
    
    
    
    PicMenu(6).Visible = Not (bClosedLot)
    PicMenu(4).Visible = bMeanValue
    
    USER_PATH = IIf(bClosedLot, USER_DATA_PATH, USER_TEMP_PATH)
    
    
    SetMeanTable (bMeanValue)

Else
End If



End Sub



Private Sub PicMenu_Click(Index As Integer)
Image3_Click Index
End Sub

Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To PicMenu.Count - 1
    If i = Index Then
        PicMenu(i).BackColor = &H505050
    Else
        PicMenu(i).BackColor = vbColorDarkFont
    End If
Next
End Sub


Private Sub Picture1_Click()
Picture1.BackColor = vbColorTextBlue ' &H8000&
DefaultMenu_Click 8

End Sub



Private Sub Text1_Change(Index As Integer)
If Me.Visible Then
    Text1(0).Top = Label1(0).Top + Label1(0).Height + Text1(0).Height / 20
    'Text1(Index).ForeColor = vbWhite
    If dIndexProcedura > 0 Then
       
        If Len(Text1(Index)) > 0 Then
            If (InStr(Text1(Index), Combo1)) Then
                Picture1.Visible = False
                DefaultMenu(8).Visible = False
            Else
                If Combo1.ListIndex = 0 Then
                    Picture1.Visible = True
                    DefaultMenu(8).Visible = True
                Else
                    Picture1.Visible = False
                    DefaultMenu(8).Visible = False
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub Text1_Click(Index As Integer)
Dim sString As String
Dim rc As Boolean

Select Case Index
    Case 0
        
        sString = Text1(0)
        If InStr(sString, Combo1) Then sString = ""
        If F_InputBox.DoShow(("Enter ") & Combo1, ("Filter : ") & Combo1, True, ("Apply"), ("Exit"), sString, Image1) Then
            Text1(0) = sString
           
            GlobalSearch
            Call ExistCampioneInTabella
            
            rc = IIf(Len(Text1(0)) > 0, True, False)
            'DefaultMenu(5).Visible = rc
            
            If dIndexProcedura > 0 Then
            
            
                If Combo1.ListIndex = 0 Then
                    Picture1.Visible = rc
                    DefaultMenu(8).Visible = rc
                Else
                    Picture1.Visible = False
                    DefaultMenu(8).Visible = False
                End If
            End If

            
            
            
            
        End If
End Select
End Sub

Private Sub Timer2_Timer()

Dim i As Integer
    '
    ' start form
    '
      bHilight = True

    If CampioneSelezionato <> "" Then
        Text1(0) = CampioneSelezionato
        ExistCampioneInTabella
    End If
     RiempiGrid GrdLot
    Timer2.Enabled = False
    
    
End Sub

Private Sub Timer3_Timer()
    If MyLbHelpCount > 1 Then
      '  lbMenuHelp(0).Visible = False
      '  lbMenuHelp(1).Visible = False
        
        Picture1.BackColor = vbColorTextDarkBlue
        MyLbHelpCount = 0
    End If
    MyLbHelpCount = MyLbHelpCount + 1
End Sub


Private Sub GetFormDatiCampione()
Dim sString As String
Dim rc As Boolean
    sString = Text1(0)
    If InStr(Text1(0), Combo1) Then sString = ""
   ' rc = F_AnagraficaCampione.DoShow(sString)
    If rc Then
        PopupMessage 2, ("Campione salvato in Archivio"), , , ("Anagrafica Campione"), Image1
    End If
                
End Sub


Private Sub ScorriTabella(ByVal bValue As Boolean)

Dim MyRow As Integer
If GrdLot.Rows > 1 Then
    MyRow = IIf(bValue, (IndexTabella * 10) + 10, (IndexTabella * 10) - 19)
    IndexTabella = IIf(bValue, IndexTabella + 1, IndexTabella - 1)
    If IndexTabella < 1 Then
        IndexTabella = 1
        GrdLot.Cell(1, 1).EnsureVisible
    ElseIf MyRow >= GrdLot.Rows Then
        GrdLot.Cell(GrdLot.Rows - 1, 1).EnsureVisible
        IndexTabella = MaxIndex
    'ElseIf IndexTabella >= MaxIndex - 1 And Not (bValue) Then
        'GrdLot.Cell((IndexTabella) * 10, 1).EnsureVisible
    Else
         
        GrdLot.Cell(MyRow, 1).EnsureVisible
    
    End If
End If

End Sub




Private Function ExistCampioneInTabella() As Boolean
    Dim NumCol As Integer
    Dim rc As Boolean
    Dim i As Integer
    
    rc = True
    If InStr(Text1(0), Combo1) Then Exit Function
    
    rc = IIf(GrdLot.Rows < 2, False, True)
    
  '  NumCol = SetNumCol(Combo1)

    
    If Not (rc) And Combo1.ListIndex = 0 Then
        'If ExistCampione(Text1(0)) = False Then
        '    If F_MsgBox.DoShow(("Campione nuovo, inserire le specifiche in Archivio?"), Text1(0), True, ("SI"), ("NO"), Image1) Then
        '        DefaultMenu_Click 7
        '    End If
            
       ' End If
    End If
  
    ExistCampioneInTabella = rc
    
    
End Function
Private Function SetNumCol(ByVal sString As String) As Integer

Select Case Trim(UCase(sString))
    Case UCase(("Code"))
        SetNumCol = 2
    Case UCase(("Description"))
        SetNumCol = 3
    Case UCase(("Lot"))
        SetNumCol = 1
    Case UCase(("Recipe"))
        SetNumCol = 4
    
    End Select

End Function
Private Function GlobalSearch()
    
    RiempiGrid GrdLot
    '

End Function

Private Function OpenIntervalloDate()



Picture2.Visible = True

    






End Function





Private Sub IntervalloDate(ByVal bValue As Boolean)
    
    
    MyPeriodo = IIf(bValue, "", MyPeriodo)
    ' se falso allora ripristino il periodo di filtro
    

    
    bFiltroIntervalloDate = bValue
    ctlCalendar1.Visible = False
    Picture2.Visible = False
     If bValue = False Then
        lbDataFiltro(0) = ""
        lbDataFiltro(1) = ""
        CmbVisual_Click
    Else
        GlobalSearch
    End If
End Sub



Public Function FillTabellaTutte(ByVal Grd As Grid, Optional ByVal Periodo As String, Optional ByVal StringaFiltro As String, Optional ByVal stringa As String, Optional MyDA As String, Optional MyA As String, Optional ByVal ChangefilterLots As Integer = 2) As Boolean
    Dim i As Integer
    Dim t As Integer
    Dim rc As Boolean
   
    Dim sName As String
    Dim sString As String
    Dim dMyDA As Date
    Dim dMyA As Date

    Dim NowRighe As Integer
    On Error GoTo ERR_FILL
    rc = True
    

    stringa = Trim(stringa)
    
    If Len(Trim(MyDA)) > 0 Then
        dMyDA = FormatDateTime(MyDA, vbShortDate)
        dMyA = FormatDateTime(MyA, vbShortDate)
    End If
            
    
    Grd.AutoRedraw = False
    Grd.Rows = 1
    
    
    If StringaFiltro = "" Then
        sString = ""
    Else
    
        If InStr(UCase(stringa), UCase(("Search"))) Then
            sString = ""
        Else
            Select Case Trim(StringaFiltro)
                Case ("Code")
                    If stringa = "" Then
                        sString = " and Code=''"
                    Else
                        sString = " and Code like '%" & Replace(Trim(stringa), "'", "''") & "%' or ReagentCode like '%" & Replace(Trim(stringa), "'", "''") & "%' or ReagentCode2 like '%" & Replace(Trim(stringa), "'", "''") & "%'"
                    End If
                Case ("Description")
                    sString = " and Description like '%" & Replace(Trim(stringa), "'", "''") & "%'"
                Case ("Lot")
                    sString = " and Lot like '%" & Replace(Trim(stringa), "'", "''") & "%' or ReagentLot like '%" & Replace(Trim(stringa), "'", "''") & "%' or ReagentLot2 like '%" & Replace(Trim(stringa), "'", "''") & "%'"
                Case ("Recipe")
                    sString = " and Recipe like '%" & Replace(Trim(stringa), "'", "''") & "%'"
            End Select
        End If
    End If
    
    Select Case ChangefilterLots
        Case 0
           ' If Len(sString) > 0 Then
                sString = " and Finished=FALSE"
           ' Else
               'sString = "Finished=FALSE"
               
               Grd.Column(21).Width = 0
           ' End If
        Case 1
            'If Len(sString) > 0 Then
                sString = " and Finished=TRUE"
            'Else
               ' sString = "Finished=TRUE"
           ' End If
           
               Grd.Column(21).Width = 100
    
    End Select
    With dbTabReport
        .filter = ""
        If Periodo <> "" Then
            Periodo = FormatDateTime(Periodo, vbShortDate)
            .filter = "StartDate>=#" & Periodo & "# " & sString
        Else
            .filter = "StartDate>=#" & dMyDA & "# AND StartDate<=#" & dMyA & "# " & sString
            
        End If
        
        If .EOF Then
            GoTo ERR_END:
        Else
            '  trovato qualcosa....
        End If
            .MoveFirst
            Do
                With Grd
                    .AddItem "", False
                    .Cell(.Rows - 1, 0).Text = .Rows - 1
                    .Cell(.Rows - 1, 1).Text = IIf(IsNull(Trim(dbTabReport!Lot)), "", "    " & Trim(dbTabReport!Lot))
                    .Cell(.Rows - 1, 2).Text = IIf(IsNull(Trim(dbTabReport!code)), "", "    " & Trim(dbTabReport!code))
                    .Cell(.Rows - 1, 3).Text = IIf(IsNull(Trim(dbTabReport!Description)), "", "    " & Trim(dbTabReport!Description))
                    .Cell(.Rows - 1, 4).Text = IIf(IsNull(Trim(dbTabReport!Recipe)), "", Trim(dbTabReport!Recipe))
                    .Cell(.Rows - 1, 5).Text = IIf(IsNull(Trim(dbTabReport!PREPWK)), "", Trim(dbTabReport!PREPWK))
                    .Cell(.Rows - 1, 6).Text = IIf(IsNull(Trim(dbTabReport!RangeMin)), "", Trim(dbTabReport!RangeMin))
                    .Cell(.Rows - 1, 7).Text = IIf(IsNull(Trim(dbTabReport!RangeMax)), "", Trim(dbTabReport!RangeMax))
                    .Cell(.Rows - 1, 8).Text = IIf(IsNull(Trim(dbTabReport!StartDate)), "", FormatDataLAT(Trim(dbTabReport!StartDate)))
                    .Cell(.Rows - 1, 9).Text = IIf(IsNull(Trim(dbTabReport!Exp)), "", Trim(dbTabReport!Exp))
                    .Cell(.Rows - 1, 10).Text = IIf(IsNull(Trim(dbTabReport!TestNumber)), "0", Trim(dbTabReport!TestNumber))
                    .Cell(.Rows - 1, 11).Text = IIf(IsNull(Trim(dbTabReport!Evaluation)), "False", Trim(dbTabReport!Evaluation))
                    .Cell(.Rows - 1, 12).Text = IIf(IsNull(Trim(dbTabReport!Finished)), "False", Trim(dbTabReport!Finished))
                    .Cell(.Rows - 1, 13).Text = IIf(IsNull(Trim(dbTabReport!Operator)), "", Trim(dbTabReport!Operator))
                    .Cell(.Rows - 1, 14).Text = IIf(IsNull(Trim(dbTabReport!Note)), "", Trim(dbTabReport!Note))
                    .Cell(.Rows - 1, 15).Text = dbTabReport!ID
                    .Cell(.Rows - 1, 16).Text = IIf(IsNull(Trim(dbTabReport!Nomefile)), "", Trim(dbTabReport!Nomefile))
                    .Cell(.Rows - 1, 17).Text = IIf(IsNull(Trim(dbTabReport!NomeFileReport)), "", Trim(dbTabReport!NomeFileReport))
                    .Cell(.Rows - 1, 18).Text = IIf(IsNull(Trim(dbTabReport!NomeFileExcel)), "", Trim(dbTabReport!NomeFileExcel))
                   
                   
                    .Cell(.Rows - 1, 22).Text = IIf(IsNull(Trim(dbTabReport!ReagentLot)), "", Trim(dbTabReport!ReagentLot))
                    .Cell(.Rows - 1, 23).Text = IIf(IsNull(Trim(dbTabReport!ReagentCode)), "", Trim(dbTabReport!ReagentCode))
                     
                    .Cell(.Rows - 1, 24).Text = IIf(IsNull(Trim(dbTabReport!ReagentLot2)), "", Trim(dbTabReport!ReagentLot2))
                    .Cell(.Rows - 1, 25).Text = IIf(IsNull(Trim(dbTabReport!ReagentCode2)), "", Trim(dbTabReport!ReagentCode2))
                   
                    For t = 1 To .Cols - 1
                        If CBool(.Cell(.Rows - 1, 12).Text) = True Then
                            .Cell(.Rows - 1, t).ForeColor = &H404040
                        Else
                             .Cell(.Rows - 1, t).ForeColor = vbBlack
                        End If
                    
                    Next
                    
                      If dbTabReport!ExcelDone Then
                        .Cell(.Rows - 1, 21).BackColor = vbColorGreen
                        .Cell(.Rows - 1, 21).Text = "OK"
                        .Cell(.Rows - 1, 21).ForeColor = vbWhite
                        
                    End If
                    
                    
                    
        '.Cell(0, 0).Text = "n."
        '.Cell(0, 1).Text = "Lot Number"
        '.Cell(0, 2).Text = "Code SFG"
        '.Cell(0, 3).Text = "Description"
        '.Cell(0, 4).Text = "Recipe"
        '.Cell(0, 5).Text = "Prep. Week"
        '.Cell(0, 6).Text = "Range Min"
        '.Cell(0, 7).Text = "Range Max"
        '.Cell(0, 8).Text = "Date"
        '.Cell(0, 9).Text = "Exp.Date"
        '.Cell(0, 10).Text = "# Test" ' quanti test ho fatto
        '.Cell(0, 11).Text = "Mean Value" ' se ho fatto calcolo medie
        '.Cell(0, 12).Text = "Finalise" ' se ho finalizzat ( solo Laboratory Manager )
        '.Cell(0, 13).Text = "QC Operator"
        '.Cell(0, 14).Text = "QC Note"
        ' .Cell(0, 15).Text = "ID"
        ' .Cell(0, 16).Text = "FileName"
        '.Cell(0, 17).Text = "NomeFileReport"
        '.Cell(0, 18).Text = "NomeFileExcel"
        '.Cell(0, 19).Text = "CODE_ID"
                    
                End With
                .MoveNext
            Loop Until .EOF
            
    End With

    
ERR_END:
    On Error GoTo 0
    For i = 1 To Grd.Rows - 1
        For t = 0 To Grd.Cols - 1
            If t = 1 Or t = 2 Or t = 3 Or t = 16 Then
                Grd.Cell(i, t).Alignment = cellLeftCenter
            Else
                Grd.Cell(i, t).Alignment = cellCenterCenter
            End If
            'Grd.Cell(i, t).ForeColor = vbColorTextDarkBlue
        Next
    Next
    Grd.Column(1).AutoFit
    Grd.Column(2).AutoFit
    Grd.Column(3).AutoFit
    Grd.Column(4).AutoFit
    Grd.Column(5).AutoFit
    Grd.Column(6).AutoFit
    Grd.Column(7).AutoFit
    Grd.Column(8).AutoFit
    Grd.Column(9).AutoFit
    Grd.Column(10).AutoFit
    Grd.Column(11).AutoFit
    Grd.Column(12).AutoFit
    
    Grd.Column(22).AutoFit
    Grd.Column(23).AutoFit
    Grd.Column(24).AutoFit
    Grd.Column(25).AutoFit
    
    Grd.AutoRedraw = True
    Grd.Refresh
    FillTabellaTutte = rc
    Exit Function
ERR_FILL:
    rc = False
    MsgBox err.Description
    Resume Next
End Function

Private Function CancellaTab() As Boolean
    
    If MyID > 0 Then
        If F_MsgBox.DoShow(("Delete Selected Record ?"), "Database", , ("Delete"), ("Exit")) Then
            
            If CancellaRecord(MyID) Then
                Text1(0) = ""
                GrdLot.ReadOnly = False
                GrdLot.Selection.DeleteByRow
                GrdLot.ReadOnly = True
                UploadDownloadMessageCounter = 0
                PopupMessage 2, ("Record Deleted..."), , , PROGRAM_NAME, Image1
                m_rc = True
            Else
            End If
        End If
    End If
End Function

Private Function ChiudiTab() As Boolean
 If MyID > 0 Then
        If F_MsgBox.DoShow(("Close Selected Lot ?"), "Database", , ("Close"), ("Exit")) Then
            
            If CloseRecord(MyID) Then
               
                Image3_Click 0
               
                UploadDownloadMessageCounter = 0
                PopupMessage 2, ("Lot Closed..."), , , PROGRAM_NAME, Image1
                m_rc = True
            Else
            End If
        End If
    End If
End Function
Private Function CloseRecord(ByVal MyID As Long) As Boolean
Dim rc As Boolean
Dim U_PATH As String
Dim sString As String

U_PATH = USER_PATH
    CloseSettingDataFile
    On Error GoTo ERR_CAN
    rc = True
    With dbTabReport
        .filter = ""
        .filter = "ID='" & MyID & "'"
        If .EOF Then
        Else
            !Finished = True
            !ClosingDate = FormatDataLAT(Now())
            If IsNull(Trim(!PREPWK)) Or Trim(!PREPWK) = "" Then
                sString = ""
                If F_InputBox.DoShow("Enter value...", "PREPARATION WEEK", , , , sString) Then
                    !PREPWK = sString
                End If
            End If
            
            
            ' Exp
            If IsNull(Trim(!Exp)) Or Trim(!Exp) = "" Then
                sString = ""
                If F_InputBox.DoShow("Enter value...", "Exp", , , , sString) Then
                    !Exp = sString
                End If
            End If
                        
            
            .Update
            DoEvents
       
            .Close
            .Open "SELECT *  FROM TabReport  order by id -1 ", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
        End If
        
        If MyFileName <> "" Then
            
            
            If FileExists(USER_TEMP_PATH & MyFileName) Then
            
            USER_PATH = USER_TEMP_PATH
            
            ' modifico il file :: closed!!!

            sString = ""
            If F_InputBox.DoShow("Enter value...", "Validation Date", , , , sString) Then
                 sString = FormatDataLAT(sString)
            End If


            SaveSettingData SettingName, "Close QC", "Date", FormatDateTime(Now, vbShortDate)
            SaveSettingData SettingName, "Close QC", "Validation Date", sString
            SaveSettingData SettingName, "Close QC", "Operator", MyOperatore.Name
            
            CloseSettingDataFile
            If USER_PATH = USER_DATA_PATH Then
                ' non c'è bisogno di spostarlo/cancellarlo----
            Else
                FileCopy USER_TEMP_PATH & MyFileName, USER_DATA_PATH & MyFileName
                Kill USER_TEMP_PATH & MyFileName
            End If
            DoEvents
            End If
        End If
    End With

ERR_END:
    On Error GoTo 0
    USER_PATH = U_PATH
    CloseSettingDataFile
    CloseRecord = rc
    Exit Function
ERR_CAN:
    rc = False
    MsgBox err.Description
    Resume Next
End Function

Private Function CancellaRecord(ByVal MyID As Long) As Boolean
Dim rc As Boolean

    On Error GoTo ERR_CAN
    rc = True
    With dbTabReport
        .filter = ""
        .filter = "ID='" & MyID & "'"
        If .EOF Then
        Else
            .Delete
            .Update
        
            ' cancello anche il file.....
            If MyFileName <> "" Then
                If FileExists(USER_TEMP_PATH & MyFileName) Then Kill USER_TEMP_PATH & MyFileName
                If FileExists(USER_DATA_PATH & MyFileName) Then Kill USER_DATA_PATH & MyFileName
            End If
        End If
    
    End With

ERR_END:
    On Error GoTo 0
    CancellaRecord = rc
    Exit Function
ERR_CAN:
    rc = False
    MsgBox err.Description
    Resume Next
End Function

Private Function SetMeanTable(ByVal bValue As Boolean) As Boolean
Dim rc As Boolean
Dim AndOr As String
On Error GoTo ERR_SET
    rc = True
    If MyFileName <> "" Then
        
    Else
        bValue = False
        rc = False
    End If
      
    AndOr = Chr$(247)
             
    If bValue Then
    
        CloseSettingDataFile
        
        GetcodeInformaton
    
        GetMeanTable Grd3, MyFileName
        Call SortGrid(Grd3)
       
          If InStr(MeasurementUnit, "mg") Then
                UNIT_PP = "ppm"
            Else
                UNIT_PP = "ppb"
            End If

            Grd3.Cell(0, 2).Text = "Target Value " & AndOr & " U [" & UNIT_PP & "]"
            Grd3.Cell(0, 3).Text = "Mean Value [" & UNIT_PP & "]"
            Grd3.Cell(0, 4).Text = "Tot Average [" & UNIT_PP & "]"
            
        CloseSettingDataFile
       ' Grd3.Cell(0, 2).Text = "Target Value " & AndOr & " U"
    
    Else
        Grd3.Rows = 1
    
    End If


ERR_END:
    On Error GoTo 0
    SetMeanTable = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox err.Description
    Resume Next
End Function

Private Sub GetcodeInformaton()
  
    With dbTabCode
        .filter = ""
        .filter = "Code='" & Trim(sCode) & "'"

           ' AndOr = Chr$(247)
       
        If .EOF Then
           Exit Sub
        Else
        
            MeasurementUnit = IIf(IsNull(Trim(!MeasurementUnit)), "", " " & Trim(!MeasurementUnit))
        End If
    End With
            
End Sub
Private Function ChangeLabelLots()
    Dim i As Integer
    For i = 0 To 2
        If i = IndexOpenClosedLot Then
            Label4(i).ForeColor = vbColorTextDarkBlue 'vbColorOrange
        Else
             Label4(i).ForeColor = vbWhite
        End If
    Next
End Function
