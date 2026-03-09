VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{93E1D7CD-F84F-4C8F-BDFF-C4C1AD9E3B89}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form F_PRINT 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chemical Production"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8250
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "F_PRINT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00886010&
      BorderStyle     =   0  'None
      Caption         =   "&H00886010&"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   1320
      Width           =   2775
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00886010&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   1215
         MouseIcon       =   "F_PRINT.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   60
         Width           =   315
      End
   End
   Begin MSComctlLib.ProgressBar ProgressWord 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Timer ProgTime 
      Interval        =   60
      Left            =   2880
      Top             =   1440
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00886010&
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
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8295
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00886010&
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
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   240
         Picture         =   "F_PRINT.frx":0316
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lbl_Print 
         BackStyle       =   0  'Transparent
         Caption         =   "....."
         ForeColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   5625
      End
      Begin VB.Label lbl_Form 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Printing Report NR..."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   1920
      Top             =   2720
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   4
      TextControl     =   0   'False
      ListBoxControl  =   0   'False
      PictureControl  =   0   'False
      FrameControl    =   0   'False
      DriveListBoxControl=   0   'False
      ListViewControl =   0   'False
      FileListBoxControl=   0   'False
      DirListBoxControl=   0   'False
   End
End
Attribute VB_Name = "F_PRINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_rc As Boolean
Private WordApp As Object
Private StartTime As Boolean
Private FinalPrint As Boolean
Private sNumReport As String
Private bError As Boolean
Private sPrint As Boolean
Public Function DoShow(ByVal mCodice As String, Optional ByVal bPrint As Boolean = True) As Boolean

    Dim m_FlgLoading As Boolean
    
    On Error GoTo ERR_SHOW
    
    m_rc = False
    
    m_FlgLoading = False
    sNumReport = mCodice
    sPrint = bPrint
    Call DoInitialize
    
    Me.Show vbModal
    
    If m_rc = True Then
            'in caso di modifiche
            
            
    End If
    
    DoShow = m_rc
    
ERR_END:
    On Error GoTo 0
    Exit Function
    

ERR_SHOW:

    m_rc = False
    Resume ERR_END
 
End Function

Private Sub Form_Load()
    Dim sString As String
    'WindowsXPC1.InitSubClassing
    If sPrint Then
        lbl_Print.Caption = ("Wait : Printing Document, press Exit to continue.")
    Else
        lbl_Print.Caption = ("Wait : Printing Document, press Exit to continue.")
    End If
    With ProgressWord
        .Min = 0
        .Max = 100
        .Value = 0
    End With
    StartTime = False

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyEscape Then
    '    Unload Me
    'End If
End Sub
Private Sub DoInitialize()
    Dim sString As String
    If sPrint = False Then
        Me.Caption = App.EXEName & " - Microsoft Word"
    End If
    sString = "Printing pdf file..."
    
 
        If SetPrinterMaterialRequisition(sNumReport) Then
            bError = False
        Else
            lbl_Print.Caption = "Errore nella crezione e stampa del file."
            bError = True
        End If
    
End Sub
Private Sub cmd_Form_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0
            DoInitialize
        Case 1
           WordApp.ActiveDocument.Close (False)
            WordApp.Quit
            Set WordApp = Nothing

            
            DoEvents
            m_rc = True
            Unload Me
    End Select
End Sub



Private Function SetPrinterMaterialRequisition(ByVal NumReport As String) As Boolean
    Dim rc As Boolean
    Dim sPath As String
    Dim sPre As String
    Dim mPath As String
    Dim sString As String
    Dim sMyLoadPath As String
    Dim strReport As String
    rc = True
    
    strReport = "MR-"

    If CreateWord(WordApp) Then
            StartTime = True
            
            sString = "" '("Rapporto n")
            
           
            sString = sString & strReport & NumReport
       
                                 
            mPath = USER_DOCUMENTI
            sPath = mPath & PathRequisition
            
            sMyLoadPath = CheckSavePath(mPath & "Bin")
            
            sString = FormatNomeFile(sString)
            
        If SettSavePath(sPath) Then
            If CreaReport(WordApp, NumReport, sString, sMyLoadPath, sPath) Then
                FinalPrint = True
            Else
                FinalPrint = False
                rc = False
            End If
        End If
    Else
        rc = False
    End If
    SetPrinterMaterialRequisition = rc
End Function

Private Sub Stampato(ByVal OkPrint As Boolean)
    Dim OkString As String
    Dim KoString As String
    If bError Then
        OkString = ("Warning : An error occured...")
        KoString = ("Warning : An error occured...")
    Else
        If sPrint = False Then
            OkString = ("Document correctly done. Exit to continue.")
            KoString = ("Warning : An error occured...")
        Else
            OkString = ("Document correctly done. Exit to continue.")
            KoString = ("Warning : An error occured...")
        End If
    End If
    
    If OkPrint Then
        lbl_Print = OkString
    Else
        lbl_Print = KoString
    End If
  '  cmd_Form(1).Enabled = True
End Sub

Private Sub Frame1_Click()
cmd_Form_Click 1
End Sub


Private Sub Label1_Click(Index As Integer)
cmd_Form_Click 1
End Sub

Private Sub ProgTime_Timer()
    On Error Resume Next
    If StartTime Then
        With ProgressWord
            If .Value < 100 Then
                .Value = .Value + 1
            Else
                StartTime = False
                Stampato (FinalPrint)
            End If
        End With
    End If
    On Error GoTo 0
End Sub
