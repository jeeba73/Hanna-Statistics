VERSION 5.00
Begin VB.Form FormQC 
   BackColor       =   &H00877773&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7230
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
   ScaleHeight     =   7230
   ScaleWidth      =   15390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   7095
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   15255
      Begin VB.ComboBox cbRegistration 
         Height          =   375
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   5400
         Visible         =   0   'False
         Width           =   3495
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
         Left            =   9000
         TabIndex        =   26
         Top             =   5400
         Width           =   3495
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
         Left            =   9000
         TabIndex        =   24
         Top             =   4920
         Width           =   3495
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
         Index           =   4
         Left            =   3360
         TabIndex        =   22
         Top             =   4920
         Width           =   3255
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
         Left            =   9000
         TabIndex        =   20
         Top             =   4440
         Width           =   3495
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
         Left            =   9000
         TabIndex        =   18
         Top             =   6240
         Width           =   3015
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Save QC"
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
            TabIndex        =   19
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
         Index           =   2
         Left            =   3360
         TabIndex        =   13
         Top             =   4440
         Width           =   3255
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
         Left            =   9000
         TabIndex        =   12
         Top             =   3960
         Width           =   3495
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
         Left            =   3360
         TabIndex        =   11
         Top             =   3960
         Width           =   3255
      End
      Begin VB.Frame frCommandInside 
         BackColor       =   &H000040C0&
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
         Height          =   975
         Index           =   1
         Left            =   2400
         TabIndex        =   9
         Top             =   1680
         Width           =   3500
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Failed"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
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
            TabIndex        =   10
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Frame frCommandInside 
         BackColor       =   &H00A88030&
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
         Height          =   975
         Index           =   0
         Left            =   6000
         TabIndex        =   7
         Top             =   1680
         Width           =   3500
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Waiting"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
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
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Frame frCommandInside 
         BackColor       =   &H00208040&
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
         Index           =   3
         Left            =   9600
         TabIndex        =   5
         Top             =   1680
         Width           =   3500
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Passed"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
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
            TabIndex        =   6
            Top             =   360
            Width           =   3495
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
         Index           =   2
         Left            =   12120
         TabIndex        =   3
         Top             =   6240
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
            Index           =   2
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
         Height          =   1455
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15255
         Begin VB.Label lbInside 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "Preparation QC"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   705
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   15060
         End
      End
      Begin VB.Label lbFormulation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correction Date"
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
         Left            =   7080
         TabIndex        =   27
         Top             =   5400
         Width           =   1785
      End
      Begin VB.Label lbFormulation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correction"
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
         Left            =   7800
         TabIndex        =   25
         Top             =   4920
         Width           =   1065
      End
      Begin VB.Label lbFormulation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QC Operator"
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
         Left            =   1800
         TabIndex        =   23
         Top             =   4920
         Width           =   1155
      End
      Begin VB.Label lbFormulation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration"
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
         Left            =   7800
         TabIndex        =   21
         Top             =   4440
         Width           =   1065
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   15000
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Label lbQC 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preparation QC"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   570
         Left            =   0
         TabIndex        =   17
         Top             =   2880
         Width           =   15255
      End
      Begin VB.Label lbFormulation 
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
         Index           =   2
         Left            =   2520
         TabIndex        =   16
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label lbFormulation 
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
         Index           =   1
         Left            =   7800
         TabIndex        =   15
         Top             =   3960
         Width           =   1065
      End
      Begin VB.Label lbFormulation 
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
         Left            =   2520
         TabIndex        =   14
         Top             =   3960
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   15120
         Y1              =   6120
         Y2              =   6120
      End
   End
End
Attribute VB_Name = "FormQC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_rc As Boolean
Private MyID As Long
Private RecipeCode As String
Private strQC As String
Private SettingFileName As String
Private RecipeQC As QCType
Private bCorrection As Boolean


Public Function DoShow(ByVal Code As String, ByVal FileName As String, ByVal ID As Long, ByRef PrepQC As String) As Boolean
Dim FormTop As Long
    On Error GoTo ERR_SHOW
    SettingFileName = FileName
    m_rc = False
    mOk
    Call SetcbRegistration
    RecipeCode = Code
    MyID = ID
    FormTop = Screen.Height / 2 - Me.Height / 2
    Me.Top = FormTop + 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    
    lbInside(0) = Code
    txFormulation(1) = MyOperatore.Name
    
    With iQRCodeType
        If .Recipe <> "" And .bCheck Then
    
            Call FillFormWithQRCode
        End If
    End With
    
    
    

    Me.Show vbModal
    
    If m_rc = True Then
        PrepQC = strQC
        
    End If
ERR_END:
    On Error GoTo 0
    DoShow = m_rc
    Exit Function
ERR_SHOW:
    m_rc = False
    Resume ERR_END
End Function

Private Function FillFormWithQRCode()
    
    
    With iQRCodeType
       
        Select Case .QC
            Case "Passed"
                frCommandInside_Click 3
            Case Else
                frCommandInside_Click 1
        End Select
        
        lbInside(0) = .Code
        txFormulation(0) = .Date
        txFormulation(2) = .Note
        txFormulation(4) = .Operator
        txFormulation(3) = .Tablet

        lbQC = "QC : " & .QC

    End With
    
    
    
End Function


Private Sub cbRegistration_Click()
txFormulation(3) = cbRegistration
cbRegistration.Visible = False

If InStr(cbRegistration, "Book") Then txFormulation_Click 3
End Sub

Private Sub frCommandInside_Click(Index As Integer)
    Select Case Index
        Case 0
            strQC = "Waiting"
            Me.BackColor = frCommandInside(Index).BackColor
            lbInside(0).BackColor = Me.BackColor
        Case 1
           ' If CheckPrivilege(1) Then
                strQC = "Failed"
                Me.BackColor = frCommandInside(Index).BackColor
                lbInside(0).BackColor = Me.BackColor
           ' End If
        Case 3
           ' If CheckPrivilege(1) Then
                txFormulation(1) = MyOperatore.Name
                strQC = "Passed"
                Me.BackColor = frCommandInside(Index).BackColor
                bCorrection = CheckIfPreparationCorrection(MyID)
                lbInside(0).BackColor = Me.BackColor
                
                If bCorrection Then
                    EnterCorrection
                End If
                
           ' End If
        Case 2
            Unload Me
        Case 4
            ' save QC
            m_rc = True
            
            If bCorrection Then
                If EnterCorrection Then
                Else
                    Exit Sub
                End If
            End If
            
            Call SaveNewQC
            
            
    End Select
     If Me.Visible Then
        lbQC = "QC : " & strQC
        txFormulation(0).SetFocus
     End If


End Sub


Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub

Private Sub txFormulation_Change(Index As Integer)
Dim rc As Boolean

rc = IIf(Len(Trim(txFormulation(Index))) > 0, True, False)
txFormulation(Index).BackColor = IIf(rc, &HF0F0F0, vbRed)

End Sub

Private Sub txFormulation_Click(Index As Integer)
Dim Answer As String
Dim Selected As String
Dim sString As String
Dim bNumber As Boolean

Selected = "Preparation"
Answer = txFormulation(Index)
sString = lbFormulation(Index)


If Index = 0 Then If Answer = "" Then Answer = FormatDataLAT(Now())
If Index = 6 Then If Answer = "" Then Answer = FormatDataLAT(Now())

If Index = 1 Then
    
    If frmLogin.DoShow Then
        txFormulation(1) = MyOperatore.Name
        Exit Sub
    Else
        Exit Sub
    End If
    
End If
        
    If Index = 3 Then
        If (InStr(txFormulation(Index), "Book")) Then
           ' GoTo cont:
        Else
           ' Call cbRegVisibile
        End If
    Else
cont:
        If F_InputBox.DoShow(sString, Selected, , , , Answer, , bNumber, Me.Top) Then
        
            txFormulation(Index) = Answer
            
            Select Case Index
                Case 0, 6
                    ' isdate?
                    If IsDate(Answer) Then
                         txFormulation(Index) = FormatDataLAT(Answer)
                         uPreparation.PreparationDate = Answer
                         
                         
                    Else
                        PopupMessage 2, "Please enter a valid Date...", , True
                    End If
            End Select
        End If
        
    End If




End Sub

Private Sub cbRegVisibile()
    cbRegistration.Left = txFormulation(3).Left
    cbRegistration.Top = txFormulation(3).Top
    cbRegistration.Width = txFormulation(3).Width
    cbRegistration.ZOrder
    cbRegistration.Visible = True
End Sub


Private Function SaveNewQC() As Boolean


    If txFormulation(0) = "" Then
        PopupMessage 2, "Please enter a valid Date...", , True
        Exit Function
    End If


    If txFormulation(1) = "" Then
        PopupMessage 2, "Please enter an Operator...", , True
        Exit Function
    End If
    
    If txFormulation(3) = "" Then
        PopupMessage 2, "Please fill Registration field...", , True
        Exit Function
    End If
    

    If strQC = "" Then
        PopupMessage 2, "Please Select QC", , True
        Exit Function
    End If
    
    If F_MsgBox.DoShow("Add QC : " & strQC & " to Recipe?", RecipeCode) = False Then Exit Function

   


    With RecipeQC
        .Date = txFormulation(0)
        .Operator = txFormulation(1)
        .Note = txFormulation(2)
        .SettingName = SettingFileName
        .RecipeCode = RecipeCode
        .Status = strQC
        .Registration = txFormulation(3)
        
        .QCOperator = txFormulation(4)
        .Correction = txFormulation(5)
        .CorrectionDate = FormatDataLAT(txFormulation(6))
        
        .ID = MyID
    End With


    Call SetQcStatus(RecipeQC)
    Call SetTabPreparationToQC(RecipeQC)

    PopupMessage 2, "QC Added..."
    Unload Me
End Function

Private Sub SetcbRegistration()
With cbRegistration
    .Clear
    .AddItem "Tablet QC n."
    .AddItem "Book n."
End With
End Sub

Private Sub txFormulation_DblClick(Index As Integer)
    If Index = 3 Then
        Call cbRegVisibile
    End If
End Sub


Private Function EnterCorrection() As Boolean

Dim rc As Boolean
Dim i As Integer

 EnterCorrection = True
For i = 5 To 6
   rc = IIf(Len(Trim(txFormulation(i))) > 0, True, False)
   txFormulation(i).BackColor = IIf(rc, &HF0F0F0, vbRed)
   If rc = False Then EnterCorrection = False
Next

    If Not (rc) Then
        PopupMessage 2, "Please fill Correction Data...", , , lbInside(0)
    End If

End Function
