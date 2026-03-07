VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form F_PICTOGRAM 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12555
   ClientLeft      =   0
   ClientTop       =   0
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12555
   ScaleMode       =   0  'User
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00473733&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1095
      Left            =   240
      TabIndex        =   18
      Top             =   1320
      Width           =   18735
      Begin VB.PictureBox Pictogram 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   6
         Left            =   13320
         Picture         =   "F_PICTOGRAM.frx":0000
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   27
         Top             =   180
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox Pictogram 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   7
         Left            =   15360
         Picture         =   "F_PICTOGRAM.frx":44F2
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   26
         Top             =   180
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox Pictogram 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   8
         Left            =   17280
         Picture         =   "F_PICTOGRAM.frx":89E4
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   25
         Top             =   180
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox Pictogram 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   3
         Left            =   6960
         Picture         =   "F_PICTOGRAM.frx":CED6
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   24
         Top             =   180
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox Pictogram 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   4
         Left            =   9120
         Picture         =   "F_PICTOGRAM.frx":113C8
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   23
         Top             =   180
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox Pictogram 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   5
         Left            =   11160
         Picture         =   "F_PICTOGRAM.frx":158BA
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   22
         Top             =   180
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox Pictogram 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   0
         Left            =   960
         Picture         =   "F_PICTOGRAM.frx":19DAC
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   21
         Top             =   180
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox Pictogram 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   1
         Left            =   2880
         Picture         =   "F_PICTOGRAM.frx":1E29E
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   20
         Top             =   180
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox Pictogram 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   2
         Left            =   4800
         Picture         =   "F_PICTOGRAM.frx":22790
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   19
         Top             =   180
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   1320
      Top             =   1200
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   240
      ScaleHeight     =   9135
      ScaleWidth      =   18735
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   18735
      Begin VB.Frame Frame2 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   9135
         Left            =   6120
         TabIndex        =   15
         Top             =   0
         Width           =   6495
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Echipamente de siguranta"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1140
            Left            =   960
            TabIndex        =   17
            Top             =   360
            Width           =   4725
            WordWrap        =   -1  'True
         End
         Begin VB.Label lbSafety 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "None"
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
            Height          =   5685
            Left            =   240
            TabIndex        =   16
            Top             =   2040
            Width           =   6000
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   9135
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   6135
         Begin VB.Label lbStatement 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Statement"
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
            Height          =   5445
            Left            =   240
            TabIndex        =   14
            Top             =   2040
            Width           =   5625
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Afirmatie"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   0
            TabIndex        =   13
            Top             =   360
            Width           =   6135
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precautie"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   570
         Left            =   14670
         TabIndex        =   11
         Top             =   360
         Width           =   2205
      End
      Begin VB.Label lbPrecaution 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "None"
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
         Height          =   5685
         Left            =   12720
         TabIndex        =   10
         Top             =   2040
         Width           =   6000
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   18000
         Picture         =   "F_PICTOGRAM.frx":26C82
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00473733&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   2640
      ScaleHeight     =   615
      ScaleWidth      =   16335
      TabIndex        =   6
      Top             =   600
      Width           =   16335
      Begin VB.Label lbCode 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hanna Code"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Index           =   1
         Left            =   -2280
         TabIndex        =   7
         Top             =   45
         Width           =   18465
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      ScaleHeight     =   615
      ScaleWidth      =   2415
      TabIndex        =   4
      Top             =   600
      Width           =   2415
      Begin VB.Label lbCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hanna Code"
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
         Height          =   345
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   120
         Width           =   1770
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
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   11760
      Width           =   18735
      Begin VB.Label lbCommandInside 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Checked"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   40
         Width           =   18735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19335
      Begin VB.Image Image7 
         Height          =   240
         Left            =   18720
         Picture         =   "F_PICTOGRAM.frx":2A064
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lbInside 
         BackStyle       =   0  'Transparent
         Caption         =   "Hazardous Statement"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   1935
         WordWrap        =   -1  'True
      End
   End
   Begin FlexCell.Grid Grid1 
      Height          =   8895
      Left            =   360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2640
      Width           =   18615
      _ExtentX        =   32835
      _ExtentY        =   15690
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
   Begin VB.Shape shInside 
      BackColor       =   &H00808080&
      BorderColor     =   &H00808080&
      Height          =   12555
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   19200
   End
End
Attribute VB_Name = "F_PICTOGRAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_rc As Boolean
Private Code As String
Private Lot As String
Private IndexClassification As Integer
Private ClassificationCodes() As Variant
Private ClassificationCodesClean As Variant
Private PictogramCodes() As String
Private userCode As String
Private RecipeCode As String
Private UserID As Long
Private bUserFormVisible As Boolean


Private Sub Frame3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Frame3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Frame3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
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

End Sub




Public Function DoShow(ByVal ID As Long, Optional ByVal Index As Integer, Optional ByVal Code As String, Optional bFormVisible As Boolean = True) As Boolean
Dim rc As Boolean
Dim strClassification As String

    On Error GoTo ERR_SHOW
    
    UserID = ID

    Call SetGrid1
    IndexClassification = Index
    userCode = Code

    ReDim ClassificationCodes(0)
    
   
    Select Case Index
        Case 0
            ' hanna code
            'With dbTabCodeClassification
            '    .filter = ""
             '   .filter = "Code='" & userCode & "'"
             '   If .EOF Then
             '       PopupMessage 2, "No hazardous statement for this Code", , , userCode
             '       GoTo ERR_END:
             '   Else
             '
             '       ID = !ID
             '   End If
            'End With
            RecipeCode = ""
            With dbTabCode
                .filter = ""
                .filter = "Code='" & userCode & "'"
                If .EOF Then
                    PopupMessage 2, "No hazardous statement for this Code", , , userCode
                   GoTo ERR_END:
                Else
             
                    RecipeCode = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
                End If
            End With
            
            
            If RecipeCode = "" Then
                PopupMessage 2, "No hazardous statement for this Code", , , userCode
                GoTo ERR_END:
            End If
            
            strClassification = ("Hanna Code: " & userCode & " | User : " & MyOperatore.Name)
            
        Case 1
            ' raw material
            
        Case 2
            ' recipe
            strClassification = ("Recipe : " & userCode & " | User : " & MyOperatore.Name)
    End Select
    
    Call CreateClassificationLogFile(strClassification)
    UserID = ID
    bUserFormVisible = bFormVisible
    
  
  
 ' Dim rc As Boolean
    rc = GetPictogram(UserID)
    Picture3.Visible = IIf(bUserFormVisible, rc, False)
    Timer1.Enabled = False
    

    Me.Show vbModal
    
    If m_rc = True Then

    End If
ERR_END:
    On Error GoTo 0
    DoShow = m_rc
    Exit Function
ERR_SHOW:
    m_rc = False
    MsgBox err.Description
    Resume ERR_END
End Function


Private Sub Form_Resize()
Picture3.Left = Me.Width / 2 - Picture3.Width / 2

End Sub






Private Sub frCommandInside_Click(Index As Integer)
Unload Me
End Sub

Private Sub Image1_Click()
'Picture3.Visible = False
End Sub

Private Sub Image7_Click()
Unload Me
End Sub




Private Function GetPictogram(ByVal RM_ID As Long) As Boolean
Dim rc As Boolean
Dim strClassificationCodes As String
Dim strPictogramCodes As String

rc = True
'If RM_ID > 0 Then
    Select Case IndexClassification
        Case 0
            '-------------------------------------
            ' pictogram per Code
            '-------------------------------------
            lbCode(1) = userCode

            With dbTabCodeClassification
                .filter = ""
                .filter = "Code='" & Replace(userCode, "'", "''") & "'"
                If .EOF Then Exit Function
                strClassificationCodes = IIf(IsNull(!Phrases), "", Trim(!Phrases))
            End With
            
            With dbTabRecipe
                .filter = ""
                .filter = "Code='" & RecipeCode & "'"
                If .EOF Then
                Else
                          lbCode(1) = userCode & "   |   Recipe : " & RecipeCode
                         strClassificationCodes = IIf(IsNull(!Classification), "", Trim(!Classification))
           
                End If
            
            End With
        Case 1
            '-------------------------------------
            ' pictogram per RawMaterial
            '-------------------------------------
            lbCode(0) = "Chemical RM"
            
            With dbTabRawMaterial
                .filter = ""
                .filter = "ID='" & RM_ID & "'"
                If .EOF Then
                    GoTo cont
                Else
                
                
                    strClassificationCodes = IIf(IsNull(!Classification), "", Trim(!Classification))
                    strPictogramCodes = IIf(IsNull(!Pictograms), "", Trim(!Pictograms))

                     lbCode(1) = GetChemicalRMbyID(RM_ID) & IIf(IsNull(dbTabRawMaterial!ChemicalReactionLiquid), "", "  -  " & Trim(dbTabRawMaterial!ChemicalReactionLiquid))
           
                End If
            End With
        Case 2
            '-------------------------------------
            ' pictogram per Recipe
            '-------------------------------------
             lbCode(0) = "Recipe"
            
            With dbTabRecipe
                .filter = ""
                .filter = "ID='" & RM_ID & "'"
                If .EOF Then
                    GoTo cont
                Else
                
                    lbCode(1) = Trim(!Code) & " - " & Trim(!Description)
                    DoEvents
                    
                    strClassificationCodes = IIf(IsNull(!Classification), "", Trim(!Classification))
           
                End If
            End With

    End Select
    
    rc = SetStringAndPictogram(PictogramCodes(), strClassificationCodes, strPictogramCodes, ClassificationCodes(), Me)
    If rc Then
        SetStatement
    End If
'End If
cont:
GetPictogram = rc

End Function




Private Sub Label3_Click()
Picture3.Visible = Not (Picture3.Visible)
End Sub

Private Sub labPrec_Click()
Picture3.Visible = Not (Picture3.Visible)
End Sub

Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub


Private Sub SetGrid1()
Dim i As Integer
With Grid1



      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25

        .Cols = 6
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize

        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Code"
        .Cell(0, 2).Text = "Statement"
        .Cell(0, 3).Text = "HazardCategory"
        .Cell(0, 4).Text = "Precaution"
        .Cell(0, 5).Text = "SafetyEquipments"
       

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 150
            .Cell(0, i).FontBold = True
            
        Next
        
        .Column(0).Width = 0
'        .Column(1).Width = 150
        .BoldFixedCell = True
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh

End With



End Sub
Private Sub SetStatement()
Dim i As Integer
Dim MaxCount As Integer
With Grid1



      .Rows = 1
        lbSafety = ""
        lbPrecaution = ""
        
        
        .AutoRedraw = False
        lbStatement = ""
        

           With dbTabFrasiH
           
           
                For i = LBound(ClassificationCodes) To UBound(ClassificationCodes)
                        .filter = ""
                        .filter = "Code='" & Trim(ClassificationCodes(i)) & "'"
                        If .EOF Then
                        
                        
                        Else
                            Grid1.AddItem "", False
                            Grid1.Cell(Grid1.Rows - 1, 1).Text = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
                            Grid1.Cell(Grid1.Rows - 1, 2).Text = IIf(IsNull(Trim(!Statement)), "", Trim(!Statement))
                            Grid1.Cell(Grid1.Rows - 1, 3).Text = IIf(IsNull(Trim(!HazardCategory)), "", Trim(!HazardCategory))
                            Grid1.Cell(Grid1.Rows - 1, 4).Text = IIf(IsNull(Trim(!Precaution)), "", Trim(!Precaution))
                            Grid1.Cell(Grid1.Rows - 1, 5).Text = IIf(IsNull(Trim(!SafetyEquipments)), "", Trim(!SafetyEquipments))
                            Grid1.Cell(Grid1.Rows - 1, 3).Alignment = cellCenterCenter
                            Grid1.Cell(Grid1.Rows - 1, 4).Alignment = cellCenterCenter
                         
                            If InStr(lbSafety, Grid1.Cell(Grid1.Rows - 1, 5).Text) Then
                            Else
                                lbSafety = setSafety(lbSafety, Grid1.Cell(Grid1.Rows - 1, 5).Text)
                            End If
                            
                            If InStr(lbPrecaution, Grid1.Cell(Grid1.Rows - 1, 4).Text) Then
                            Else
                                lbPrecaution = setSafety(lbPrecaution, Grid1.Cell(Grid1.Rows - 1, 4).Text)
                            End If
                            
                            
                            If InStr(lbStatement, Grid1.Cell(Grid1.Rows - 1, 2).Text) Then
                            Else
                                lbStatement = setSafety(lbStatement, Grid1.Cell(Grid1.Rows - 1, 2).Text)
                            End If
                            
                           
                        End If
                
                
                Next
        
               ' .Cell(0, 1).Text = "Code"
               ' .Cell(0, 2).Text = "Statement"
               ' .Cell(0, 3).Text = "HazardCategory"
               ' .Cell(0, 4).Text = "Precaution"
               ' .Cell(0, 5).Text = "SafetyEquipments"
           
           End With

        '.Column(1).AutoFit
        .Column(2).AutoFit
        '.Column(3).AutoFit
       ' .Column(4).AutoFit
        .Column(5).AutoFit
        
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh

End With



End Sub




Private Sub lbSafety_Click()
'Picture3.Visible = False
End Sub

Private Sub lbPrecaution_Click()
'Picture3.Visible = False
End Sub

Private Function setSafety(ByVal strSafety As String, ByVal strNew As String) As String

Dim vettore As Variant
Dim i As Integer

 vettore = Split(strNew, "-")
If InStr(strNew, Chr$(10)) Then
    'vettore = Split(strNew, Chr$(13))
    vettore = Split(strNew, Chr$(10))
End If

'MsgBox strNew

    For i = LBound(vettore) To UBound(vettore)
    
        If InStr(UCase(strSafety), Trim(UCase(vettore(i)))) Then
        Else
            strSafety = strSafety & vbCrLf & Trim(vettore(i))
        End If
        
    Next
    

setSafety = strSafety


End Function

Private Sub lbPrecaution_Change()
lbPrecaution.FontSize = IIf(Len(lbPrecaution) > 235, 12, 20)
End Sub

Private Sub lbSafety_Change()
lbSafety.FontSize = IIf(Len(lbSafety) > 235, 12, 20)
End Sub

Private Sub lbStatement_Change()
lbStatement.FontSize = IIf(Len(lbStatement) > 235, 12, 20)
End Sub



Private Sub Timer1_Timer()
'Dim rc As Boolean
   ' rc = GetPictogram(UserID)
    'Picture3.Visible = IIf(bUserFormVisible, rc, False)
    Timer1.Enabled = False
End Sub
