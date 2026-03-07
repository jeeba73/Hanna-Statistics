VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form FormCodes 
   BackColor       =   &H00886010&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15390
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleMode       =   0  'User
   ScaleWidth      =   15390
   ShowInTaskbar   =   0   'False
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
      Height          =   8850
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   15255
      Begin VB.Frame frCommandInside 
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
         Index           =   2
         Left            =   5880
         TabIndex        =   15
         Top             =   8160
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
            Index           =   2
            Left            =   0
            TabIndex        =   16
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.ComboBox cmbLine 
         Height          =   375
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   8160
         Width           =   2775
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
         Left            =   9000
         TabIndex        =   8
         Top             =   8160
         Visible         =   0   'False
         Width           =   3015
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Select"
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
            TabIndex        =   9
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
         Left            =   5160
         TabIndex        =   6
         Top             =   3600
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
            Index           =   1
            Left            =   1920
            TabIndex        =   7
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
         Index           =   1
         Left            =   12120
         TabIndex        =   4
         Top             =   8160
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
            Index           =   1
            Left            =   0
            TabIndex        =   5
            Top             =   120
            Width           =   3015
         End
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
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   15255
         Begin VB.TextBox txCode 
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
            Height          =   520
            Left            =   5880
            TabIndex        =   13
            Top             =   0
            Width           =   3495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Search Code"
            Height          =   375
            Left            =   4440
            TabIndex        =   14
            Top             =   105
            Width           =   1335
         End
         Begin VB.Image Image1 
            Height          =   360
            Left            =   9720
            Picture         =   "FormCodes.frx":0000
            Top             =   60
            Width           =   360
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   120
            X2              =   15120
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Database"
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
            Left            =   14160
            TabIndex        =   3
            Top             =   240
            Width           =   885
         End
         Begin VB.Label lbInside 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hanna FG Codes Database "
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
            Left            =   240
            TabIndex        =   2
            Top             =   120
            Width           =   3210
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00E0E0E0&
            X1              =   120
            X2              =   15120
            Y1              =   480
            Y2              =   480
         End
      End
      Begin FlexCell.Grid Grid1 
         Height          =   7335
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   720
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   12938
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   8160
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   15120
         Y1              =   5400
         Y2              =   5400
      End
   End
End
Attribute VB_Name = "FormCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyID As Long
Private m_rc As Boolean
Private uHannaFGCode As String
Private indexProcedure As Integer
Private MyRecipe As String
Public Function DoShow(Optional ByRef UserHannaFGCode As String, Optional FormTop As Double, Optional ByRef ID As Long, Optional ByVal Index = 0, Optional ByRef strRecipe As String) As Boolean


    On Error GoTo ERR_SHOW
    indexProcedure = Index
   
    
    m_rc = False
    mOk
    If FormTop = 0 Then FormTop = Screen.Height / 2 - Me.Height / 2
    Me.Top = FormTop + 300
    Me.Left = Screen.Width / 2 - Me.Width / 2

    Call SetCodeGrid(Grid1)
    Call SetLine(cmbLine, True)

    Me.Show vbModal
    
    
    If m_rc = True Then
        UserHannaFGCode = uHannaFGCode
        strRecipe = MyRecipe
        ID = MyID
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





Private Sub cmbLine_Click()
InsertCode Grid1
End Sub

Private Sub frCommandInside_Click(Index As Integer)
Select Case Index
    Case 0
        m_rc = IIf(MyID > 0, True, False)
    Case 1
        m_rc = False
    Case 2
       ' Call F_PICTOGRAM.DoShow(MyID, 0, uHannaFGCode)
        Exit Sub
End Select

Unload Me


End Sub

Private Sub Grid1_DblClick()
If MyID > 0 Then frCommandInside_Click IIf(indexProcedure = 0, 0, 2)
End Sub

Private Sub Grid1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

MyID = 0
frCommandInside(0).Visible = False
frCommandInside(2).Visible = False
MyRecipe = ""
If FirstRow > 0 Then
    MyID = Grid1.Cell(FirstRow, 3).Text
    uHannaFGCode = Trim(Grid1.Cell(FirstRow, 1).Text)
    frCommandInside(0).Visible = True
    frCommandInside(2).Visible = True
   ' MyRecipe = Trim(Grid1.Cell(FirstRow, 4).Text)
End If



End Sub

Private Sub Image1_Click()
txCode = ""
End Sub

Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub



Private Sub SetCodeGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
      .Rows = 3

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        
        
        .Cols = 4
        .RowHeight(0) = 35
        
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Hanna FGCode"
        .Cell(0, 2).Text = "Product Name"
       ' .Cell(0, 3).Text = "Line"
       ' .Cell(0, 4).Text = "Recipe"
       ' .Cell(0, 5).Text = "Mix #1"
       ' .Cell(0, 6).Text = "Mix #2"
        .Cell(0, 3).Text = "ID"
     
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 130
            .Cell(0, i).FontBold = True
        Next
        
        .Column(3).Width = 0
        .Column(2).Width = 250

        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
    Call InsertCode(Grd)

End Sub


Private Sub InsertCode(ByVal Grd As Grid)
Dim i As Integer
Dim sString As String
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        With dbTabFinishGood
            .filter = ""
            
            If txCode <> "" Then sString = "Code like '*" & Trim(txCode) & "*'"
            If InStr(LCase(cmbLine), "all lines") Then
             
            Else
               ' If sString <> "" Then sString = sString & " and "
               ' sString = sString & " line='" & cmbLine & "'"
            End If
           
           .filter = sString
           
           If .EOF Then
           
           Else
           
                .MoveFirst
           
                For i = 1 To .RecordCount
                    Grd.AddItem "", False
                    
        
                    Grd.Cell(Grd.Rows - 1, 0).Text = i
                    Grd.Cell(Grd.Rows - 1, 1).Text = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
                    Grd.Cell(Grd.Rows - 1, 2).Text = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
                    Grd.Cell(Grd.Rows - 1, 3).Text = !ID
                    'Grd.Cell(Grd.Rows - 1, 3).Text = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
                    'Grd.Cell(Grd.Rows - 1, 4).Text = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
                    'Grd.Cell(Grd.Rows - 1, 5).Text = IIf(IsNull(Trim(!Mix1)), "", Trim(!Mix1))
                    'Grd.Cell(Grd.Rows - 1, 6).Text = IIf(IsNull(Trim(!Mix2)), "", Trim(!Mix2))
                    
                    .MoveNext
                Next
           End If
        
        
        End With
        
        
        Dim t As Integer
        For t = 1 To .Rows - 1
            
            For i = 1 To .Cols - 1
                .Column(i).Alignment = IIf(i > 2, cellCenterCenter, cellLeftCenter)
            Next
       Next
     
        '.Column(1).AutoFit
        '.Column(4).AutoFit
        '.Column(5).AutoFit
        '.Column(6).AutoFit
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
    

End Sub

Private Sub txCode_Change()
   SearchCode txCode, Grid1, False
    Dim rc As Boolean
    rc = IIf(Len(Trim(txCode)) > 0, True, False)
    txCode.BackColor = IIf(rc, vbWhite, &HF0F0F0)
End Sub
