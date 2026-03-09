VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form F_PICTOGRAM 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13290
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
   ScaleHeight     =   8610
   ScaleWidth      =   13290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   240
      Top             =   720
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   360
      ScaleHeight     =   5175
      ScaleWidth      =   12615
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   12615
      Begin VB.Label lbSafety 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   0
         TabIndex        =   20
         Top             =   1560
         Width           =   12615
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   12240
         Picture         =   "F_PICTOGRAM.frx":0000
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Safety Equipments"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   480
         Width           =   12615
      End
   End
   Begin VB.PictureBox Pictogram 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   8
      Left            =   8040
      Picture         =   "F_PICTOGRAM.frx":0A02
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Pictogram 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   7
      Left            =   7080
      Picture         =   "F_PICTOGRAM.frx":4EF4
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Pictogram 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   6
      Left            =   6120
      Picture         =   "F_PICTOGRAM.frx":93E6
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Pictogram 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   5
      Left            =   5160
      Picture         =   "F_PICTOGRAM.frx":D8D8
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Pictogram 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   4
      Left            =   4200
      Picture         =   "F_PICTOGRAM.frx":11DCA
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Pictogram 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   3
      Left            =   3240
      Picture         =   "F_PICTOGRAM.frx":162BC
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Pictogram 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   2
      Left            =   2280
      Picture         =   "F_PICTOGRAM.frx":1A7AE
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Pictogram 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   1
      Left            =   1320
      Picture         =   "F_PICTOGRAM.frx":1ECA0
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Pictogram 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   0
      Left            =   360
      Picture         =   "F_PICTOGRAM.frx":23192
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00473733&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2640
      ScaleHeight     =   495
      ScaleWidth      =   10335
      TabIndex        =   6
      Top             =   840
      Width           =   10335
      Begin VB.Label lbCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hanna Code"
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
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   90
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   2295
      TabIndex        =   4
      Top             =   840
      Width           =   2295
      Begin VB.Label lbCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hanna Code"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   1260
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
      Left            =   360
      TabIndex        =   2
      Top             =   7920
      Width           =   12615
      Begin VB.Label lbCommandInside 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Checked"
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
         TabIndex        =   3
         Top             =   120
         Width           =   12615
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
      Width           =   13335
      Begin VB.Image Image7 
         Height          =   240
         Left            =   12840
         Picture         =   "F_PICTOGRAM.frx":27684
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
      Height          =   4935
      Left            =   480
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2640
      Width           =   12375
      _ExtentX        =   21828
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Safety Equipments"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00644603&
      Height          =   270
      Left            =   10905
      TabIndex        =   21
      Top             =   1920
      Width           =   1995
   End
   Begin VB.Shape shInside 
      BackColor       =   &H00808080&
      BorderColor     =   &H00808080&
      Height          =   8610
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   13275
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
Private UserHannaCode As String
Private UserID As Long
Private bUserFormVisible As Boolean

Private Sub Frame3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Form_MouseDown Button, Shift, x, y
End Sub

Private Sub Frame3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Form_MouseMove Button, Shift, x, y
End Sub

Private Sub Frame3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
FrmMove = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMove = True
    DragX = x
    DragY = y
    If Me.WindowState = 2 Then
        FrmMove = False
       
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nx, ny
    If Me.WindowState = 2 Then
        FrmMove = False
        Exit Sub
    End If
    nx = Me.Left + x - DragX
    ny = Me.Top + y - DragY
    Me.Left = nx
    Me.Top = ny
    FrmMove = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer



If Me.WindowState = 2 Then
    FrmMove = False
End If
Dim nx, ny
    If FrmMove Then
        nx = Me.Left + x - DragX
        ny = Me.Top + y - DragY
        Me.Left = nx
        Me.Top = ny
    End If

End Sub




Public Function DoShow(ByVal ID As Long, Optional ByVal Index As Integer, Optional ByVal HannaCode As String, Optional bFormVisible As Boolean = True) As Boolean
Dim rc As Boolean
    On Error GoTo ERR_SHOW
    

    Call SetGrid1
    IndexClassification = Index
    UserHannaCode = HannaCode

    ReDim ClassificationCodes(0)
    
    If Index = 0 Then
        ' hanna code
        With dbTabCodeClassification
            .filter = ""
            .filter = "Code='" & UserHannaCode & "'"
            If .EOF Then
                PopupMessage 2, "No hazardous statement for this Code", , , UserHannaCode
                GoTo ERR_END:
            Else
            
                ID = !ID
            End If
        End With
    
    End If
    UserID = ID
    bUserFormVisible = bFormVisible
    
  

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


Private Sub Form_Resize()
Picture3.Left = Me.Width / 2 - Picture3.Width / 2

End Sub






Private Sub frCommandInside_Click(Index As Integer)
Unload Me
End Sub

Private Sub Image1_Click()
Picture3.Visible = False
End Sub

Private Sub Image7_Click()
Unload Me
End Sub




Private Function GetPictogram(ByVal MR_ID As Long) As Boolean
Dim rc As Boolean
Dim strClassificationCodes As String
Dim strPictogramCodes As String

rc = True
If MR_ID > 0 Then
    Select Case IndexClassification
        Case 0
            '-------------------------------------
            ' pictogram per Code
            '-------------------------------------
            lbCode(1) = UserHannaCode
            With dbTabCodeClassification
                strClassificationCodes = IIf(IsNull(!Phrases), "", Trim(!Phrases))
            End With
        Case 1
            '-------------------------------------
            ' pictogram per RawMaterial
            '-------------------------------------
            lbCode(0) = "Chemical MR"
            lbCode(1) = GetMRbyID(MR_ID) '& IIf(IsNull(dbTabMR!ChemicalReactionLiquid), "", "  -  " & Trim(dbTabMR!ChemicalReactionLiquid))
            
            With dbTabMR
                .filter = ""
                .filter = "ID='" & MR_ID & "'"
                If .EOF Then
                    GoTo cont
                Else
                
                
                    strClassificationCodes = IIf(IsNull(!Classification), "", Trim(!Classification))
                    strPictogramCodes = "" 'getPictograms(strClassificationCodes) ' IIf(IsNull(!Pictograms), "", Trim(!Pictograms))

                    
                End If
            End With
            

    End Select
    
    rc = SetStringAndPictogram(PictogramCodes(), strClassificationCodes, strPictogramCodes, ClassificationCodes(), Me)
    If rc Then
        SetStatement
    End If
End If
cont:
GetPictogram = rc

End Function




Private Sub Label3_Click()
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
        .AutoRedraw = False
        

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
                            'Grid1.Cell(Grid1.Rows - 1, 5).Alignment = cellCenterCenter
                            If InStr(lbSafety, Grid1.Cell(Grid1.Rows - 1, 5).Text) Then
                            Else
                                lbSafety = setSafety(lbSafety, Grid1.Cell(Grid1.Rows - 1, 5).Text)
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
Picture3.Visible = False
End Sub

Private Function setSafety(ByVal strSafety As String, ByVal strNew As String) As String

Dim vettore As Variant
Dim i As Integer

vettore = Split(strNew, "-")


    For i = LBound(vettore) To UBound(vettore)
    
        If InStr(strSafety, Trim(vettore(i))) Then
        Else
            strSafety = strSafety & vbCrLf & Trim(vettore(i))
        End If
        
    Next
    

setSafety = strSafety


End Function




Private Sub Timer1_Timer()
Dim rc As Boolean
    rc = GetPictogram(UserID)
    Picture3.Visible = IIf(bUserFormVisible, rc, False)
    Timer1.Enabled = False
End Sub
