VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form FormHannaCode 
   BackColor       =   &H00964901&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16815
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   16815
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
      Height          =   8055
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   16695
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
         Index           =   2
         Left            =   7320
         TabIndex        =   12
         Top             =   7440
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
            TabIndex        =   13
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
         Left            =   10440
         TabIndex        =   9
         Top             =   7440
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
            TabIndex        =   10
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
         Left            =   5760
         TabIndex        =   5
         Top             =   3120
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
            TabIndex        =   6
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
         Left            =   13560
         TabIndex        =   3
         Top             =   7440
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
            TabIndex        =   4
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
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   120
            X2              =   15120
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lbInside 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hanna Code Database "
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
            Width           =   2745
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00E0E0E0&
            X1              =   120
            X2              =   15120
            Y1              =   600
            Y2              =   600
         End
      End
      Begin FlexCell.Grid Grid1 
         Height          =   6495
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   720
         Width           =   16455
         _ExtentX        =   29025
         _ExtentY        =   11456
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
         DefaultFontSize =   9.75
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
      Begin VB.TextBox txCode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   7440
         Width           =   2415
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   15120
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Code"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   7485
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FormHannaCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private UserSTDNumber As String
Private m_rc As Boolean
Private MyID As Long

Private MRCode As String

Public Function DoShow(Optional ByRef STDNumber As String, Optional FormTop As Double, Optional ByVal rCode As String) As Boolean
Dim rc As Boolean
    On Error GoTo ERR_SHOW
    
    m_rc = False
    mOk
    If FormTop = 0 Then FormTop = Screen.Height / 2 - Me.Height / 2
    Me.Top = FormTop + 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    
    
    MRCode = rCode
   ' lbInside(0) = MRCode
    Call SetCodeGrid(Grid1)
   
    Me.Show vbModal
    
    

    
    If m_rc = True Then
    
        STDNumber = UserSTDNumber
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
    Case 0
        m_rc = IIf(Len(UserSTDNumber) > 0, True, False)
    Case 1
        m_rc = False
    Case 2
        Call F_PICTOGRAM.DoShow(MyID, 1)
        Exit Sub
       
End Select

Unload Me


End Sub

Private Sub Grid1_DblClick()
If MyID > 0 Then frCommandInside_Click 0
End Sub

Private Sub Grid1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

MyID = 0

frCommandInside(0).Visible = False
frCommandInside(2).Visible = False
If FirstRow > 0 Then
    MyID = Grid1.Cell(FirstRow, 7).Text
    UserSTDNumber = Grid1.Cell(FirstRow, 1).Text
    frCommandInside(0).Visible = True
    frCommandInside(2).Visible = True
End If



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
      MyID = 0
      
      .Rows = 3

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        
        
        .Cols = 9
        .RowHeight(0) = 35
        
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Code"
        .Cell(0, 2).Text = "Description"
        .Cell(0, 3).Text = "Location"
        .Cell(0, 4).Text = "STOCK_QTY"
        .Cell(0, 5).Text = ""
        .Cell(0, 6).Text = ""
        .Cell(0, 7).Text = "ID"
        
        
     
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 130
            .Cell(0, i).FontBold = True
        Next
        
      
        .Column(2).Width = 450
        .Column(5).Width = 40
        .Column(6).Width = 10
        .Column(7).Width = 0
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
    Call InsertCode(Grd)

End Sub

Private Sub InsertCode(ByVal Grd As Grid)
Dim i As Integer
Dim t As Integer
Dim rc As Boolean
Dim sString As String
Dim strCode As String
Dim strLocation As String

Dim MRStockQTY As Double
Dim MRMinQTY As Double
Dim usrColor As OLE_COLOR
  
  On Error GoTo ERR_INSERT:

    If MRCode <> "" Then
        
     
    End If

    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
      
        With dbTabCode
            .filter = ""
           If txCode <> "" Then sString = "Code like '*" & Trim(txCode) & "*'"
           
           .filter = sString
           
           
           If .EOF Then
           
           Else
           
                .MoveFirst
           
                For i = 1 To .RecordCount
                
                    strCode = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
                    strLocation = IIf(IsNull(Trim(!Location)), "", Trim(!Location))
                    If strLocation = "" Then strLocation = GetLocation(strCode)
                    
                
                    Grd.AddItem "", False
                    Grd.Cell(Grd.Rows - 1, 0).Text = i
                    Grd.Cell(Grd.Rows - 1, 1).Text = strCode
                    Grd.Cell(Grd.Rows - 1, 2).Text = IIf(IsNull(Trim(!ProductName)), "", Trim(!ProductName))
                    Grd.Cell(Grd.Rows - 1, 3).Text = strLocation
                    Grd.Cell(Grd.Rows - 1, 4).Text = IIf(IsNull(Trim(!STOCK_QTY)), "0", Trim(!STOCK_QTY))
                    Grd.Cell(Grd.Rows - 1, 5).Text = IIf(IsNull(Trim(!STOCK_UNIT)), "", Trim(!STOCK_UNIT))
                    Grd.Cell(Grd.Rows - 1, 7).Text = !ID
                    

                    
                        MRStockQTY = IIf(IsNull(Trim(!STOCK_QTY)) Or Trim(!STOCK_QTY) = "", 0, Trim(!STOCK_QTY))
                        MRMinQTY = IIf(IsNull(Trim(!MinQty)) Or Trim(!MinQty) = "", 0, Val(!MinQty))
                        
                        Select Case MRStockQTY - MRMinQTY
                            Case Is < 0
                                usrColor = vbColorRed
                            Case 0
                                usrColor = vbColorOrange
                            Case Is > 0
                                usrColor = vbColorGreen
                        
                        End Select
    
                  
                            Grd.Cell(Grd.Rows - 1, 1).FontBold = True
                            Grd.Cell(Grd.Rows - 1, 1).ForeColor = &H644603
                            Grd.Cell(Grd.Rows - 1, 3).FontBold = True
                            Grd.Cell(Grd.Rows - 1, 3).ForeColor = &H644603
                            Grd.Cell(Grd.Rows - 1, 5).FontBold = True
                            Grd.Cell(Grd.Rows - 1, 5).ForeColor = vbColorBlueProgram
                            Grd.Cell(Grd.Rows - 1, 5).Alignment = cellLeftCenter
                            Grd.Cell(Grd.Rows - 1, 4).FontBold = True
                            Grd.Cell(Grd.Rows - 1, 4).ForeColor = vbColorBlueProgram
                            Grd.Cell(Grd.Rows - 1, 4).Alignment = cellRightCenter
                        
                    'Next
                    
                            Grd.Cell(Grd.Rows - 1, 6).BackColor = usrColor
                    
                    
                    
                    
                    .MoveNext
                Next
           End If
        
        
        End With
        
        
        
      ' For t = 1 To .Rows - 1
            
          '  For i = 1 To .Cols - 1
               ' .Column(i).Alignment = IIf(i > 2, cellCenterCenter, cellLeftCenter)
           ' Next
      ' Next
       .Column(4).AutoFit
    
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
    
    
ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_INSERT:
    MsgBox Err.Description
    Resume Next

End Sub

Private Sub txCode_Change()
    InsertCode Grid1
End Sub
