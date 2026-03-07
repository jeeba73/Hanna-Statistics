VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form FormChemicalRM 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15390
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
   ScaleHeight     =   8970
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
      Height          =   8849
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   15280
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
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   15735
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
            Left            =   5640
            TabIndex        =   0
            Top             =   20
            Width           =   3495
         End
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
            Left            =   12000
            TabIndex        =   14
            Top             =   80
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.Image Image1 
            Height          =   360
            Left            =   9480
            Picture         =   "FormChemicalRM.frx":0000
            Top             =   80
            Width           =   360
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Search Code"
            Height          =   375
            Left            =   4200
            TabIndex        =   15
            Top             =   120
            Width           =   1335
         End
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
            Caption         =   "Chemicals RM Database "
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
            TabIndex        =   3
            Top             =   120
            Width           =   4815
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00E0E0E0&
            X1              =   120
            X2              =   15120
            Y1              =   600
            Y2              =   600
         End
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
         Left            =   240
         TabIndex        =   13
         Top             =   8280
         Width           =   1215
      End
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
         TabIndex        =   11
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
            TabIndex        =   12
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
         Left            =   5280
         TabIndex        =   6
         Top             =   3960
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
      Begin FlexCell.Grid Grid1 
         Height          =   7215
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   840
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   12726
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
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   15120
         Y1              =   5400
         Y2              =   5400
      End
   End
End
Attribute VB_Name = "FormChemicalRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private UserChCode As String
Private m_rc As Boolean
Private MyID As Long
Private indexProcedure As Integer
Private RecipeCode As String

Public Function DoShow(Optional ByRef CHCode As String, Optional FormTop As Double, Optional ByVal rCode As String, Optional ByVal Index = 0) As Boolean
Dim rc As Boolean
    On Error GoTo ERR_SHOW
    indexProcedure = Index
  
    
    m_rc = False
    mOk
    
    If FormTop = 0 Then FormTop = Screen.Height / 2 - Me.Height / 2
    Me.Top = FormTop + 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    
    
    
    RecipeCode = rCode
    lbInside(0) = RecipeCode & " : Components"
    Call SetCodeGrid(Grid1)
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

Private Sub frCommandInside_Click(Index As Integer)
Select Case Index
    Case 0
        m_rc = IIf(Len(UserChCode) > 0, True, False)
    Case 1
        m_rc = False
    Case 2
        Call F_PICTOGRAM.DoShow(MyID, 1, UserChCode)
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
If FirstRow > 0 Then
    MyID = Grid1.Cell(FirstRow, 6).Text
    UserChCode = Grid1.Cell(FirstRow, 1).Text
    frCommandInside(0).Visible = True
    frCommandInside(2).Visible = True
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
                    Grd.Cell(Grd.Rows - 1, 4).Text = IIf(IsNull(Trim(!Um)), "", Trim(!Um))
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
                    Grd.Cell(Grd.Rows - 1, 4).Text = IIf(IsNull(Trim(!Um)), "", Trim(!Um))
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
    InsertCode Grid1
    Dim rc As Boolean
    rc = IIf(Len(Trim(txCode)) > 0, True, False)
    txCode.BackColor = IIf(rc, vbWhite, &HF0F0F0)
End Sub
