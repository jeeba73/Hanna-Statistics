VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form F_MACHINE 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "&H00808080&"
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   14595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Show all Lines"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   21
      Top             =   6240
      Width           =   2655
   End
   Begin VB.ComboBox cmbLine 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox TxtCodici 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5520
      TabIndex        =   17
      Top             =   1680
      Width           =   4815
   End
   Begin VB.TextBox TxtCodici 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   15
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Frame frCommandInside 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   7320
      TabIndex        =   12
      Top             =   6240
      Width           =   2295
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
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Frame frCommandInside 
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   9720
      TabIndex        =   10
      Top             =   6240
      Width           =   2295
      Begin VB.Label lbCommandInside 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
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
         TabIndex        =   11
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Frame frCommandInside 
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   12120
      TabIndex        =   8
      Top             =   6240
      Width           =   2295
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
         Index           =   2
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.TextBox TxtCodici 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   10560
      TabIndex        =   4
      Top             =   1680
      Width           =   3855
   End
   Begin VB.ComboBox cmbCodici 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Text            =   "Machine name"
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox TxtCodici 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   2
      Top             =   960
      Width           =   6855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16695
      Begin VB.Label lbInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Machine Database"
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
         Left            =   600
         TabIndex        =   1
         Top             =   120
         Width           =   6135
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   1
         Left            =   120
         Picture         =   "F_MACHINE.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   240
      End
   End
   Begin FlexCell.Grid GrdCodici 
      Height          =   3735
      Left            =   600
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2280
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   6588
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   10560
      TabIndex        =   20
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   18
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   16
      Top             =   1440
      Width           =   600
   End
   Begin VB.Shape shInside 
      BackColor       =   &H00808080&
      BorderColor     =   &H00808080&
      Height          =   6885
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   14595
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Heads Number"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   10560
      TabIndex        =   7
      Top             =   1440
      Width           =   1440
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   6
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Machine"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   720
      Width           =   825
   End
End
Attribute VB_Name = "F_MACHINE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_rc As Boolean
Private m_index As Integer
Private m_Codici As String
Private OldIndex As Integer
Private bSingola As Boolean
Private bNextControllo As Boolean

Private MyID As Integer
Private bAllLines As Boolean
Public Function DoShow(Optional ByRef MyName As String, Optional ByRef bControlo As Boolean = False) As Boolean
    
    On Error GoTo ERR_SHOW
    
    m_rc = False
    
    m_Codici = MyName
    
    Call SetTabellaMachine(GrdCodici)
    
    Call FillTabella
    
    Call GetUserLine
    Call SetCmbLine
    Call fillcmbMachine(cmbCodici)
    
    
    If m_Codici <> "" Then
        cmbCodici = m_Codici
        'bControlo = bNextControllo
    End If

    
    Me.Show vbModal
    
    If m_rc = True Then
        MyName = m_Codici
        bControlo = bNextControllo
    End If
ERR_END:
    On Error GoTo 0
    DoShow = m_rc
    Exit Function
ERR_SHOW:
    m_rc = False
    Resume ERR_END
End Function


Private Sub SetCmbLine()

    Call SetLine(cmbLine, True)

    If UserLine <> "" And UserLineIndex > 0 Then
        
        cmbLine = UserLine
       
    
    End If
    
End Sub


Private Sub Check1_Click()
Dim rc As Boolean
rc = IIf(Check1.Value = 1, True, False)
bAllLines = rc
Check1.FontBold = rc
Check1.ForeColor = IIf(rc, vbColorOrange, vbBlack)
Call ClearTxt
Call FillTabella
Call fillcmbMachine(cmbCodici)
End Sub

Private Sub cmbCodici_Change()
Call ClearTxt
End Sub

Private Sub cmbCodici_Click()
Call ClearTxt
Call FillTxt(cmbCodici)
End Sub

Private Sub Form_Activate()
    
    DropShadow Me.hWnd

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            m_rc = False
            Unload Me
    End Select
End Sub
Private Sub ClearTxt()
    
    Dim i As Integer
    For i = 0 To TxtCodici.Count - 1
        TxtCodici(i) = ""
    Next
End Sub
Private Function FillTabella() As Boolean
    Dim rc As Boolean
    rc = GetCodici
    FillTabella = rc
End Function

Private Function GetCodici(Optional ByVal MyName As String) As Boolean
    Dim rc As Boolean
    Dim i As Integer
      
    On Error GoTo ERR_GET
    rc = True
    
    With GrdCodici
     
        With dbTabMachine
            .filter = ""
            
            If UserLine <> "" And bAllLines = False Then
                .filter = "Line='" & UserLine & "'"
            End If
            
            If MyName <> "" Then
                 .filter = "MACHINE='" & Replace(MyName, "'", "''") & "'"
            End If
            
            If .EOF Then
               
                GrdCodici.Rows = 1
             Else
                .MoveFirst
                GrdCodici.Rows = .RecordCount + 1
                For i = 1 To .RecordCount
                    
Insert:
                    GrdCodici.Cell(i, 1).Text = IIf(IsNull(!Machine), "", Trim(!Machine))
                    GrdCodici.Cell(i, 2).Text = IIf(IsNull(!Model), "", Trim(!Model))
                    GrdCodici.Cell(i, 3).Text = IIf(IsNull(!SerialNumber), "", Trim(!SerialNumber))
                    GrdCodici.Cell(i, 4).Text = IIf(IsNull(!Description), "", Trim(!Description))
                    GrdCodici.Cell(i, 5).Text = IIf(IsNull(!HEADS_NUMBER), "", Trim(!HEADS_NUMBER))
                    GrdCodici.Cell(i, 6).Text = IIf(IsNull(!Line), "", Trim(!Line))
                    .MoveNext
                Next
            End If
        End With
    End With
    
ERR_END:
    On Error GoTo 0
    GetCodici = rc
    Exit Function
ERR_GET:
    rc = False
   ' Call TabCodiciUPDATE
    Resume Insert
End Function

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


Private Sub cmd_Archivio_Click(Index As Integer)
    Dim rc As Boolean
    
    Select Case Index
        Case 0
            '-------------------
            '  Esci
            '-------------------
            Unload Me
        Case 2

        Case 3
            '---------------------------
            '  Cancella Codici
            '----------------------------
            
            If cmbCodici <> "" Then

                m_rc = CancellaCodici
                Call ClearTxt
                Call FillTabella
                Call fillcmbMachine(cmbCodici)
            End If
        Case 8
            '-------------------
            '  Salva Codici
            '-------------------
            If cmbCodici <> "" Then
                m_rc = SalvaDatiCodici
              
                If m_rc Then
                    Call ClearTxt
                    Call FillTabella
                End If
            End If
    End Select
    

    
End Sub



Private Function SalvaDatiCodici() As Boolean
    Dim rc As Boolean
    On Error GoTo ERR_SAVE
    rc = True
   
    With dbTabMachine
        .filter = ""
        .filter = "MACHINE='" & Replace(cmbCodici, "'", "''") & "'"
        If .EOF Then .AddNew
SAVE:
        !Machine = cmbCodici
        !Description = Trim(TxtCodici(0))
        !HEADS_NUMBER = Trim(TxtCodici(1))
        !Model = Trim(TxtCodici(2))
        !SerialNumber = Trim(TxtCodici(3))
        !Line = cmbLine
        .Update
    End With
    
ERR_END:
    On Error GoTo 0
    SalvaDatiCodici = rc
    Exit Function
ERR_SAVE:
    rc = False
    MsgBox err.Description
    
    Call TabCodiciUPDATE
      
   
    Resume SAVE:
    'Resume ERR_END
    

End Function

Private Function CancellaCodici() As Boolean
    Dim rc As Boolean
    Dim i As Integer
    On Error GoTo ERR_SAVE
    rc = True

        '-----------------------------
        ' cancello la Sede
        '-----------------------------
        If F_MsgBox.DoShow("Delete Machine ?", cmbCodici) Then
        Else
            rc = False
            GoTo ERR_END
        End If
        
        With dbTabMachine
            .filter = ""
            
            .filter = "MACHINE='" & Replace(cmbCodici, "'", "''") & "'"
            If .EOF Then
            Else
                .MoveFirst
                For i = 1 To .RecordCount
                    .Delete
                    .MoveNext
                Next
            End If
        End With
        
ERR_END:
    On Error GoTo 0
    CancellaCodici = rc
    Exit Function
ERR_SAVE:
    rc = False
    MsgBox err.Description
    Resume ERR_END

End Function







Private Sub GrdCodici_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
 Dim i As Integer
    With GrdCodici
        cmbCodici = .Cell(FirstRow, 1).Text
        TxtCodici(2) = .Cell(FirstRow, 2).Text
        TxtCodici(3) = .Cell(FirstRow, 3).Text
        
        TxtCodici(0) = .Cell(FirstRow, 4).Text
        TxtCodici(1) = .Cell(FirstRow, 5).Text
        cmbLine = .Cell(FirstRow, 6).Text
        
    End With
 

End Sub


Private Sub TxtCodici_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Len(TxtCodici(Index)) > 1 And Index < TxtCodici.Count - 1 Then
      TxtCodici(Index + 1).SetFocus
    End If
End Sub



Private Function FillTxt(ByVal MyName As String)
On Error GoTo ERR_FILL:
    With dbTabMachine
        .filter = ""
        If MyName <> "" Then
            .filter = "MACHINE='" & Replace(MyName, "'", "''") & "' "
        End If
        If .EOF Then
           
         Else
ERR_SET:
                TxtCodici(0) = IIf(IsNull(!Description), "", Trim(!Description))
                TxtCodici(1) = IIf(IsNull(!HEADS_NUMBER), "", Trim(!HEADS_NUMBER))
                
                TxtCodici(2) = IIf(IsNull(!Model), "", Trim(!Model))
                TxtCodici(3) = IIf(IsNull(!SerialNumber), "", Trim(!SerialNumber))
                
                cmbLine = IIf(IsNull(!Line), UserLine, Trim(!Line))
               
        End If
    End With
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_FILL:
    Call TabCodiciUPDATE
    Resume ERR_SET:
End Function





Private Function TabCodiciUPDATE()
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '\\\\ aggiungo i campi mancanti
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    On Local Error Resume Next
   ' dbTabMachine.Update
   ' dbTabMachine.Close
   ' dbWeighCheck.Execute ("ALTER TABLE TabCodici ADD Piva char(50)WITH COMPRESSION, Telefono char(50)WITH COMPRESSION ")
   ' dbTabMachine.Open "SELECT *  FROM TabCodici", dbWeighCheck, adOpenKeyset, adLockOptimistic, adCmdText
    On Error GoTo 0
End Function


Public Function SetTabellaMachine(ByVal Grd As Grid)
       '------------------------------------------------
        '       SET TABELLA Machine
        '------------------------------------------------
Dim i As Integer

    With Grd
 
         '.Clear (True)
      .Rows = 1
      '.Redraw = False
      .DefaultRowHeight = 25
      
        '
        .AllowUserPaste = cellTextOnly
        .AllowUserResizing = False
        .ExtendLastCol = True
        .BoldFixedCell = False
        .DisplayDateTimeMask = True
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        
        .DrawMode = cellOwnerDraw
        
        .Appearance = Flat
        .ScrollBarStyle = Flat
        .FixedRowColStyle = Flat
        
        '.BackColorFixed = RGB(90, 158, 214)
        .BackColorFixedSel = RGB(110, 180, 230)
        '.BackColorBkg = RGB(90, 158, 214)
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        '.GridColor =
        '.CellBorderColorFixed = vbBlack
        

        .ButtonLocked = True
        
        .Cols = 7
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 30
        .Cell(0, 1).Text = "Code"
        .Column(1).Width = 100
        .Cell(0, 2).Text = "Model"
        .Column(2).Width = 100
        .Cell(0, 3).Text = "Serial Number"
        .Column(3).Width = 100
        .Cell(0, 4).Text = "Description"
        .Column(4).Width = 260
        .Cell(0, 5).Text = "Head Number"
        .Column(5).Width = 80
        .Cell(0, 6).Text = "Line"
        .Column(6).Width = 80
        .ReadOnly = True
        
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
        Next
        
      
   End With
 End Function
 



Public Function fillcmbMachine(ByVal cmb As ComboBox) As Boolean
Dim i As Integer
Dim rc As Boolean
    
    On Error GoTo ERR_FILL
    rc = True
    With cmb
        .Clear
        dbTabMachine.filter = ""
        If bAllLines = False And UserLine <> "" Then
            dbTabMachine.filter = "Line='" & UserLine & "'"
        End If
        If dbTabMachine.EOF Then
            rc = False
        Else
            dbTabMachine.MoveFirst
            For i = 1 To dbTabMachine.RecordCount
                .AddItem Trim(dbTabMachine!Machine)
                dbTabMachine.MoveNext
            Next
        End If
        
    End With
    
ERR_END:
    On Error GoTo 0
    fillcmbMachine = rc
    Exit Function
ERR_FILL:
    rc = False
    Resume ERR_END:
End Function



Private Sub frCommandInside_Click(Index As Integer)
Label1_Click Index
End Sub

Private Sub lbCommandInside_Click(Index As Integer)
Label1_Click Index
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
    Case 0
        cmd_Archivio_Click 8
    Case 1
        cmd_Archivio_Click 3
    Case 2
        cmd_Archivio_Click 0
End Select
End Sub


