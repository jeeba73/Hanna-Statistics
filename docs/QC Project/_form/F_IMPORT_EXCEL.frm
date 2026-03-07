VERSION 5.00
Begin VB.Form F_IMPORT_EXCEL 
   BackColor       =   &H00644603&
   BorderStyle     =   0  'None
   ClientHeight    =   7455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12330
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "F_IMPORT_EXCEL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   12330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00644603&
      Caption         =   "Delete previous records"
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
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   6180
      Width           =   4695
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00964901&
      Height          =   3405
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   11655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   3
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11040
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00964901&
      Height          =   405
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2040
      Width           =   10575
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Import"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   6600
      Width           =   9855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Import Code Database (Excel format)"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   12135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXCEL file Path :"
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
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   1875
   End
   Begin VB.Label lbIntro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This function can Import Hanna Code form an Excel File... Select file with Search , and Import."
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
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   9555
   End
End
Attribute VB_Name = "F_IMPORT_EXCEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private the_array() As String
Private num_rows As Long
Private num_cols As Long
Private m_rc As Boolean

'variabile oggetto che contiene il riferimento alla cartella di lavoro di Excel
Dim FileExcel As Object

'variabile oggetto che contiene il riferimento al foglio di lavoro di Excel
Dim FoglioExcel As Object

'variabile oggetto che contiene il  riferimento alle celle del foglio di lavoro
'di Excel
Dim CellaFoglioExcel As Range

Private bDeletePreviousRecords As Boolean


Public Function DoShow() As Boolean


    
    On Error GoTo ERR_SHOW
    List1.Clear
    m_rc = False
    
    Text3 = GetSetting(App.Title, "IMPORT CODICI", "PERCORSO", Text3)
    If Text3 <> "" Then cmdGo.Enabled = True

    
    Me.Show vbModal
    
    If m_rc = True Then
        
        

    End If
    
    DoShow = m_rc
    
ERR_END:
    On Error GoTo 0
    Exit Function
    

ERR_SHOW:

    m_rc = False
    Resume ERR_END
 
End Function

Private Sub Check1_Click()
Dim rc As Boolean
rc = IIf(Check1.Value = 1, True, False)
Check1.ForeColor = IIf(rc, vbColorOrange, vbWhite)
bDeletePreviousRecords = rc

End Sub

Private Sub cmdGo_Click()
Dim file_name As String
Dim OldCode As String
Dim NewCode As String
Dim MaxCount As Long
Dim MyBilance As Variant
Dim bCreoRichiesta As Boolean
Dim r As Long

On Error GoTo ERR_CREATE_OBJECT
    List1.AddItem Now & " - Inizio importazione da file..."
    
    cmdGo.Enabled = False
    
    If HannaCodeExcelImport(Text3, F_IMPORT_EXCEL, MaxCount, bDeletePreviousRecords) Then
        
      
    End If
    
    

END_FN:
    On Error GoTo 0
    cmdGo.Enabled = False

    Exit Sub
ERR_CREATE_OBJECT:
    MsgBox err.Description
    m_rc = False
    Resume Next
End Sub



Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
    List1.Clear
    Dim szFilename As String
    Dim PATH As String
    PATH = GetSetting(App.Title, "ImportExcel", "Path", PATH)
    szFilename = DialogFile(Me.hWnd, 1, "Open", "*.xlsx", "EXCEL Files" & Chr(0) & "*.xlsx" & Chr(0) & "Tutti i files" & Chr(0) & "*.*", PATH, "xlsx")
    
   ' szFilename = BrowseFolder(Me.hWnd, "Seleziona la cartella desiderata:")
    
    If szFilename = "" Then
       ' cmdGo.Enabled = False
        Exit Sub
    Else
        cmdGo.Enabled = True
    End If
   
    DoEvents
    Text3 = szFilename
   
End Sub


Public Function VerifyFile(FileName As String)
On Error Resume Next
Open FileName For Input As #1
If err Then
VerifyFile = False
Exit Function
End If
Close #1
VerifyFile = True
End Function



'alla chiusura del programma.......
Private Sub Form_Unload(Cancel As Integer)
Set FileExcel = Nothing 'libero ("scarico") la variabile

End Sub
