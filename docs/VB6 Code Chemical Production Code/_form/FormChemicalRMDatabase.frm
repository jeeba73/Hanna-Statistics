VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form FormChemicalRMDatabase 
   BackColor       =   &H00808080&
   Caption         =   "Chemical RM"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19005
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormChemicalRMDatabase.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12000
   ScaleMode       =   0  'User
   ScaleWidth      =   19005
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Height          =   11385
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   18855
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   10800
         Width           =   8895
         Begin VB.TextBox txCode 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   2760
            TabIndex        =   16
            Top             =   40
            Width           =   3735
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Recipes"
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
            Height          =   255
            Left            =   6720
            TabIndex        =   15
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search Code"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   105
            Width           =   2340
         End
      End
      Begin FlexCell.Grid Grid2 
         Height          =   10695
         Left            =   3720
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   -1080
         Visible         =   0   'False
         Width           =   16695
         _ExtentX        =   29448
         _ExtentY        =   18865
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColor1      =   14737632
         BackColor2      =   14737632
         BackColorBkg    =   14737632
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
         SelectionMode   =   1
         MultiSelect     =   0   'False
         DateFormat      =   0
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
         Index           =   2
         Left            =   12600
         TabIndex        =   10
         Top             =   10800
         Visible         =   0   'False
         Width           =   3015
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Acquisition History"
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
            TabIndex        =   11
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
         Left            =   9480
         TabIndex        =   7
         Top             =   10800
         Visible         =   0   'False
         Width           =   3015
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Export To Excel"
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
            TabIndex        =   8
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
         Left            =   7080
         TabIndex        =   5
         Top             =   4800
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
         Left            =   15720
         TabIndex        =   3
         Top             =   10800
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
         Width           =   18735
         Begin VB.CheckBox chAll 
            BackColor       =   &H00E0E0E0&
            Caption         =   "All Chemicals"
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
            Height          =   375
            Left            =   16200
            TabIndex        =   12
            Top             =   80
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   120
            X2              =   18480
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lbInside 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chemicals RM Database "
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00644603&
            Height          =   420
            Index           =   0
            Left            =   240
            TabIndex        =   2
            Top             =   120
            Width           =   4440
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
         Height          =   9855
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   840
         Width           =   18615
         _ExtentX        =   32835
         _ExtentY        =   17383
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
         GridColor       =   16777215
         Rows            =   1
         ScrollBarStyle  =   0
         SelectionMode   =   3
         MultiSelect     =   0   'False
         DateFormat      =   0
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
Attribute VB_Name = "FormChemicalRMDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private UserChCode As String
Private UserChCodeDescription As String
Private m_rc As Boolean
Private MyID As Long

Private RecipeCode As String



Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type

Private m_ControlGridFontSize As Double
Private m_ControlGridRowHeight As Double
Private m_ControlGridColWidth As Double
Private m_ControlPositions() As ControlPositionType
Private m_FormWid As Single
Private m_FormHgt As Single


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
        ElseIf TypeOf ctl Is Timer Then
        ElseIf TypeOf ctl Is Inet Then
        ElseIf TypeOf ctl Is Image Then
            ctl.Left = (x_scale * .Left) + IIf(x_scale = 1, 0, (x_scale - 1) * 200)
            ctl.Top = y_scale * .Top
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


Public Function DoShow(Optional ByRef CHCode As String, Optional FormTop As Double, Optional ByVal rCode As String) As Boolean
Dim rc As Boolean
    On Error GoTo ERR_SHOW
    
    m_rc = False
    mOk
    If FormTop = 0 Then FormTop = Screen.Height / 2 - Me.Height / 2
    Me.Top = FormTop + 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    
    
    RecipeCode = rCode
    lbInside(0) = RecipeCode & " : Components"
    Call SetCodeGrid(Grid1)
    Call SetChemicalRMAcquisition(Grid2)
    
    
    rc = IIf(RecipeCode <> "", True, False)
    chAll.Visible = rc
    chAll.Value = IIf(RecipeCode <> "", 0, 1)
    
     
 
    Me.Show vbModal
    
    

    
    If m_rc = True Then
    
        CHCode = UserChCode
    End If
ERR_END:
    On Error GoTo 0
    DoShow = m_rc
    Exit Function
ERR_SHOW:
    m_rc = False
    Resume ERR_END
End Function



Private Sub chAll_Click()
If chAll.Value = 1 Then
    lbInside(0) = "Chemical RM Database "
Else
    lbInside(0) = RecipeCode & " : Components"
End If
InsertCode Grid1
End Sub

Private Sub Check1_Click()
InsertCode Grid1
End Sub

Private Sub Form_Initialize()
SaveSizes
End Sub

Private Sub Form_Resize()

ResizeControls

Grid2.Move Grid1.Left, Grid1.Top, Grid1.Width, Grid1.Height
End Sub

Private Sub frCommandInside_Click(Index As Integer)


    Select Case Index
        Case 0
            ' export to excel....
            Grid2.ExportToExcel USER_DESKTOP & "\" & UserChCode & "_Acquisitions.xls", True, True
            MessageInfoTime = 2500
            PopupMessage 2, "File correcly created on Desktop", , , UserChCode & "_Acquisitions.xls"
        Case 1
            If Grid2.Visible Then
                ViewAcquisition
            Else
                Unload Me
           End If
        Case 2
            Call ViewAcquisition
    End Select

End Sub



Private Sub Grid1_DblClick()
If MyID > 0 Then frCommandInside_Click 2
End Sub

Private Sub Grid1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

MyID = 0

frCommandInside(0).Visible = False
frCommandInside(2).Visible = False
If FirstRow > 0 Then
    MyID = Grid1.Cell(FirstRow, 6).Text
    UserChCode = Grid1.Cell(FirstRow, 1).Text
    UserChCodeDescription = Grid1.Cell(FirstRow, 2).Text
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
        
        
        .Cols = 7
        .RowHeight(0) = 35
        
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Code"
        .Cell(0, 2).Text = "Description"
        .Cell(0, 3).Text = "Cas"
        .Cell(0, 4).Text = "Um"
        .Cell(0, 5).Text = "Recipe"
        .Cell(0, 6).Text = "ID"
     
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 130
            .Cell(0, i).FontBold = True
        Next
        
        .Column(5).CellType = cellCheckBox
        .Column(2).Width = 450
        .Column(5).Width = 80
        .Column(6).Width = 0
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
    Call InsertCode(Grd)

End Sub
Private Sub InsertRecipeComponent(ByVal Grd As Grid, ByVal rc As Boolean)
Dim sString As String
Dim i As Integer
Dim t As Integer

  With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        

        With dbTabRMxRecipe
            .filter = ""
            
            sString = "RecipeCode='" & RecipeCode & "' "
            
            
           If txCode <> "" Then
                sString = sString & "and CHCode like '*" & Trim(txCode) & "*'"
           End If
            
            
           
           .filter = sString
           
           
           If .EOF Then
           
           Else
           
                .MoveFirst
           
                For i = 1 To .RecordCount
                    Grd.AddItem "", False

                    Grd.Cell(Grd.Rows - 1, 0).Text = i
                    Grd.Cell(Grd.Rows - 1, 1).Text = IIf(IsNull(Trim(!CHCode)), "", Trim(!CHCode))
                    Grd.Cell(Grd.Rows - 1, 2).Text = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
                    Grd.Cell(Grd.Rows - 1, 3).Text = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
                    Grd.Cell(Grd.Rows - 1, 4).Text = IIf(IsNull(Trim(!um)), "", Trim(!um))
                    Grd.Cell(Grd.Rows - 1, 5).Text = !bMix
                    Grd.Cell(Grd.Rows - 1, 6).Text = !ID
                    
                    
                    
                    For t = 1 To Grd.Cols - 1
    
                        If !bMix Then
                            Grd.Cell(Grd.Rows - 1, t).FontBold = True
                            Grd.Cell(Grd.Rows - 1, t).ForeColor = &H644603
                        End If
                        
                    Next
                        
                    
                    
                    
                    .MoveNext
                Next
           End If
        
        
        End With
        
        

        For t = 1 To .Rows - 1
            
            For i = 1 To .Cols - 1
                .Column(i).Alignment = IIf(i > 2, cellCenterCenter, cellLeftCenter)
            Next
       Next
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With

End Sub

Private Sub InsertCode(ByVal Grd As Grid)
Dim i As Integer
Dim t As Integer
Dim rc As Boolean
Dim sString As String

    rc = IIf(Check1.Value = 1, True, False)


    If RecipeCode <> "" Then
        
        If chAll.Value = 1 Then
        Else
            ' visualizza solo quelli della ricetta...
            Call InsertRecipeComponent(Grd, rc)
            Exit Sub
        End If
    
    End If

    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        If rc Then sString = "bMix=true"
        With dbTabRawMaterial
            .filter = ""
           If txCode <> "" Then sString = "Code like '*" & Trim(txCode) & "*'" & IIf(sString <> "", " and " & sString, "")
           
           .filter = sString
           
           
           If .EOF Then
           
           Else
           
                .MoveFirst
           
                For i = 1 To .RecordCount
                    Grd.AddItem "", False
                    
        '.Cell(0, 2).Text = "Product Name"
        '.Cell(0, 3).Text = "Line"
        '.Cell(0, 4).Text = "Volume/Weight"
        '.Cell(0, 5).Text = "um"
        '.Cell(0, 6).Text = "Q.ty to produce"
        '.Cell(0, 7).Text = "Recipe"
        '.Cell(0, 8).Text = "Mix #1"
        'p.Cell(0, 9).Text = "Mix #2"
                    Grd.Cell(Grd.Rows - 1, 0).Text = i
                    Grd.Cell(Grd.Rows - 1, 1).Text = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
                    Grd.Cell(Grd.Rows - 1, 2).Text = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
                    Grd.Cell(Grd.Rows - 1, 3).Text = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
                    Grd.Cell(Grd.Rows - 1, 4).Text = IIf(IsNull(Trim(!um)), "", Trim(!um))
                    Grd.Cell(Grd.Rows - 1, 5).Text = !bMix
                    Grd.Cell(Grd.Rows - 1, 6).Text = !ID
                    
                    For t = 1 To Grd.Cols - 1
    
                        If !bMix Then
                            Grd.Cell(Grd.Rows - 1, t).FontBold = True
                            Grd.Cell(Grd.Rows - 1, t).ForeColor = &H644603
                        End If
                        
                    Next
                    
                    
                    
                    
                    .MoveNext
                Next
           End If
        
        
        End With
        
        
        
        For t = 1 To .Rows - 1
            
            For i = 1 To .Cols - 1
                .Column(i).Alignment = IIf(i > 2, cellCenterCenter, cellLeftCenter)
            Next
       Next
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
    

End Sub

Private Sub txCode_Change()
    If Grid2.Visible Then
        Call SetAcquisition(Grid2, UserChCode)
    Else
    
        InsertCode Grid1
    
    End If
End Sub


Private Sub ViewAcquisition()
Dim rc As Boolean

Grid2.Move Grid1.Left, Grid1.Top, Grid1.Width, Grid1.Height

If Grid2.Visible Then
    Label2 = "Search Code"
    lbInside(0) = "Chemicals RM Database"
    Grid2.Visible = False
    Check1.Visible = True
    frCommandInside(2).Visible = True
    frCommandInside(0).Visible = False
Else
    
    
    rc = SetAcquisition(Grid2, UserChCode)
    If rc Then
       Check1.Visible = False
       
        Label2 = "Search Recipe"
        lbInside(0) = UserChCode & " : " & UserChCodeDescription
        Grid2.ZOrder
        Grid2.Visible = True
        frCommandInside(0).Visible = True
        frCommandInside(2).Visible = True
    
    End If

End If


End Sub
Private Function SetAcquisition(ByRef Grid2 As Grid, ByVal UserChCode As String)
Dim i As Integer
Dim X As Integer
Dim PreparationID As Long
Dim PreparationLot As String
Dim t As Integer
Dim sString As String


Dim rc As Boolean
Grid2.Rows = 1

rc = True

sString = "Code='" & Trim(Replace(UserChCode, "'", "''")) & "'"
With dbTabAcquisition
    .Close
    .Open "SELECT *  FROM TabAcquisition order by AcquisitionTime ", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
    
    .filter = ""
    
    
    
     If txCode <> "" Then sString = "Code like '*" & Trim(txCode) & "*'" & IIf(sString <> "", " and " & sString, "")
   
   
    .filter = sString
    
    
    
           
    
    
    
    
    If .EOF Then
        PopupMessage 2, "No acquisition in database", , , UserChCode
        rc = False
        
    Else
        .MoveLast
         Grid2.AutoRedraw = False
         
        Grid2.AddItem "", False
        X = Grid2.Rows - 1
        
        Grid2.RowHeight(X) = 40
        Grid2.Cell(X, 1).Text = UserChCode
        Grid2.Cell(X, 2).Text = GetChemicalDescription(UserChCode)
        Grid2.Cell(X, 6).Text = GetChemicalCAS(UserChCode)
        Grid2.Cell(X, 1).FontBold = True
        Grid2.Cell(X, 2).FontBold = True
        Grid2.Cell(X, 6).FontBold = True
        
        Grid2.Cell(X, 1).FontSize = 12
        Grid2.Cell(X, 2).FontSize = 12
        Grid2.Cell(X, 6).FontSize = 12
        
         
        For i = 1 To .RecordCount
    
            Grid2.AddItem "", False
            X = Grid2.Rows - 1
            
            
            PreparationID = IIf(IsNull(Trim(!PreparationID)), 0, Trim(!PreparationID)) '!PreparationID
    
            Grid2.Cell(X, 1).Text = IIf(IsNull(Trim(!RecipeCode)), "", Trim(!RecipeCode))
            Grid2.Cell(X, 2).Text = GetRecipeDescription(Grid2.Cell(X, 1).Text)
            
           
            Grid2.Cell(X, 3).Text = IIf(IsNull(Trim(!PrepWeek)), "", Trim(!PrepWeek))
            Grid2.Cell(X, 4).Text = IIf(IsNull(Trim(!NumberPrepWeek)), "", Trim(!NumberPrepWeek))
            Grid2.Cell(X, 5).Text = IIf(IsNull(Trim(!AcquisitionTime)), "", Trim(!AcquisitionTime))
            Grid2.Cell(X, 6).Text = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
            Grid2.Cell(X, 7).Text = PadString(IIf(IsNull(Trim(!ActualWeight)), "", Trim(!ActualWeight)))
            Grid2.Cell(X, 8).Text = IIf(IsNull(Trim(!Manufacturer)), "", Trim(!Manufacturer))
            Grid2.Cell(X, 9).Text = IIf(IsNull(Trim(!ManufacturerCode)), "", Trim(!ManufacturerCode))
            Grid2.Cell(X, 10).Text = IIf(IsNull(Trim(!ManufacturerLot)), "", Trim(!ManufacturerLot))
            Grid2.Cell(X, 11).Text = IIf(IsNull(Trim(!DeliveryDate)), "", Trim(!DeliveryDate))
            Grid2.Cell(X, 12).Text = IIf(IsNull(Trim(!QtyDelivered)), "", Trim(!QtyDelivered))
            Grid2.Cell(X, 13).Text = IIf(IsNull(Trim(!WeekDelPackageNumber)), "", Trim(!WeekDelPackageNumber))
            Grid2.Cell(X, 14).Text = IIf(IsNull(Trim(!Package)), "", Trim(!Package))
            Grid2.Cell(X, 15).Text = IIf(IsNull(Trim(!Note)), "", Trim(!Note))
            Grid2.Cell(X, 16).Text = IIf(IsNull(Trim(!Operator)), "", Trim(!Operator))
            
            Grid2.Cell(X, 17).Text = IIf(IsNull(Trim(!ExpDate)), "", Trim(!ExpDate))
            
          

            With dbTabPreparation
                .filter = ""
                .filter = "ID='" & PreparationID & "'"
                If .EOF Then
                Else
                    PreparationLot = IIf(IsNull(Trim(!Lot)), "", Trim(!Lot))
                End If
            End With
            
            Grid2.Cell(X, 18).Text = PreparationLot


            Grid2.Cell(X, 7).BackColor = vbColorResults
            
            Grid2.Cell(X, 7).Alignment = cellRightCenter
            
            
            Grid2.Cell(X, 18).BackColor = vbColorResults
            'Grid2.Cell(X, 19).BackColor = vbColorResults
            'Grid2.Cell(X, 20).BackColor = vbColorResults
            'Grid2.Cell(X, 21).BackColor = vbColorResults
            
            If !bRecalculation Then
                'Grid2.Cell(x, 0).BackColor = &HC0&
                For t = 1 To 3
                    Grid2.Cell(X, t).ForeColor = &HC0&
                    Grid2.Cell(X, t).FontBold = True
                Next
            End If
            If !bRecipeComponent = False Then
               ' Grid2.Cell(x, 0).BackColor = &HFFFF&
                For t = 1 To 3
                
                    Grid2.Cell(X, t).BackColor = &HC0C0&    ' &HFFFF&
                    Grid2.Cell(X, t).ForeColor = vbWhite
                    
                    Grid2.Cell(X, t).FontBold = True
                   ' Grid2.Cell(x, t).BackColor = &HC0C0C0
                Next
            End If

            .MovePrevious
        Next
        Grid2.Column(0).Width = 0
        Grid2.Column(1).Alignment = cellLeftCenter
        Grid2.Column(2).Alignment = cellLeftCenter
        Grid2.Column(2).AutoFit
        Grid2.Refresh
        Grid2.AutoRedraw = True
        
    End If
    




End With

   SetAcquisition = rc


End Function
Public Sub SetChemicalRMAcquisition(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 19
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Recipe"
        .Cell(0, 2).Text = "Description"
        .Cell(0, 3).Text = "Prep. Week"
        .Cell(0, 4).Text = "# Prep. Week"
        .Cell(0, 5).Text = "Acquisition Time"
        
        .Cell(0, 6).Text = "CAS"
        .Cell(0, 7).Text = "Real Weight (g)"
        .Cell(0, 8).Text = "Manufacturer"
        .Cell(0, 9).Text = "Manufacturer Code"
        .Cell(0, 10).Text = "Manufacturer Lot"
        .Cell(0, 11).Text = "Delivery Date"
        .Cell(0, 12).Text = "Qty Delivered"
        .Cell(0, 13).Text = "Week Delivery"
        .Cell(0, 14).Text = "Package"
        
        .Cell(0, 15).Text = "Note"
        .Cell(0, 16).Text = "Operator"

        .Cell(0, 17).Text = "ExpDate"

        '.Cell(0, 18).Text = "Recipe"
        '.Cell(0, 19).Text = "Preparation Week"
       '.Cell(0, 20).Text = "# Prep Week"
        .Cell(0, 18).Text = "LOT"
        
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 100
            .Cell(0, i).FontBold = True
            
        Next
        
    
        .Column(4).Width = 150
        .Column(5).Width = 150
        .Column(6).Width = 150
        .Column(7).Width = 150
        .Column(8).Width = 150
        .Column(9).Width = 150
        
        .Column(12).Width = 200
        .Column(13).Width = 150
        .Column(14).Width = 150
        .Column(15).Width = 150
        .Column(16).Width = 150
        
        .Column(18).Width = 150
       ' .Column(19).Width = 150
       ' .Column(20).Width = 150
        '.Column(21).Width = 150
       '
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next

        
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With

End Sub
