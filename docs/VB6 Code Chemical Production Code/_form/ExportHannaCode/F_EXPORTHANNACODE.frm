VERSION 5.00
Begin VB.Form F_EXPORTHANNACODE 
   BackColor       =   &H00473733&
   BorderStyle     =   0  'None
   Caption         =   "Cerca / Imposta Archivio"
   ClientHeight    =   4920
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   16455
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   4920
   ScaleWidth      =   16455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      Caption         =   "Percorso file"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   720
      TabIndex        =   5
      Top             =   2160
      Width           =   15015
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   13440
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txt_ARCHIVIO 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   405
         Left            =   240
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   360
         Width           =   13095
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      ForeColor       =   &H00644603&
      Height          =   1215
      Left            =   720
      ScaleHeight     =   1215
      ScaleWidth      =   15015
      TabIndex        =   2
      Top             =   600
      Width           =   15015
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   240
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Click on ... and open browser folder. please select : dbCode.mdb"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   525
         Left            =   960
         TabIndex        =   8
         Top             =   600
         Width           =   13890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Export Hanna Code to dbCode file"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   960
         TabIndex        =   4
         Top             =   120
         Width           =   6180
      End
   End
   Begin VB.CommandButton CMD_BUTTON 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   11760
      TabIndex        =   1
      Top             =   3720
      Width           =   3975
   End
   Begin VB.CommandButton CMD_BUTTON 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   7800
      TabIndex        =   0
      Top             =   3720
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   16455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C87034&
      Height          =   4920
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   16455
   End
End
Attribute VB_Name = "F_EXPORTHANNACODE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_path As String
Private m_name As String
Private m_old_name As String
Private m_rc As Boolean
Private MyBackupPath As String
Private MyBackupName As String




Public Function DoShow(Optional ByRef t_path As String, Optional ByRef t_nome As String, Optional Index As Integer = 0) As Boolean

    Dim m_FlgLoading As Boolean
    
    On Error GoTo ERR_SHOW
    
    
    m_rc = False
    
    If Index = 0 Then
  
        MyBackupPath = t_path
        MyBackupName = dbCodeName
        
            m_path = MyBackupPath
            m_name = MyBackupName
        
    Else
            m_path = App.Path
            m_name = dbCodeName
    End If

    txt_ARCHIVIO = m_path & m_name
    
    m_FlgLoading = False
    
    Me.Show vbModal
    
    If m_rc = True Then
            'in caso di modifiche
        t_path = m_path
    End If
    
    DoShow = m_rc
    
ERR_END:
    On Error GoTo 0
    Exit Function
    

ERR_SHOW:

    m_rc = False
    Resume ERR_END
 
End Function


Private Sub CMD_BUTTON_Click(Index As Integer)
On Error GoTo ERR_CMD:

    Select Case Index
        Case 0
            
            If m_CreateArchivio(m_path, m_name, , , , , True) = True Then
                'MsgBox "Access: il file selezionato è corretto."
                m_rc = True
                If dbCode.State Then dbCode.Close
                FileCopy m_path & m_name, dbPath & dbCodeName
                
               
                
                If m_CreateArchivio(m_path, m_name) = True Then
                    PopupMessage 2, "Import Database: Operation correctly done..."
                Else
                    
                    
                    
                End If
                
                 SaveSetting App.Title, "Classification", "SetAllClassificationByRecipe", False
                 
            Else
            
                If m_CreateArchivio(dbPath, dbCodeName) = True Then
                   ' PopupMessage 2, "Old Database opened."
                End If
                    
                If F_MsgBox.DoShow(("Warning: Wrong file. Search again?"), ("Database"), False) Then
                    m_name = m_old_name
                    Command1_Click
                    
                Else
                    
                    m_rc = False
                End If
            End If
    
        Case 1
            m_rc = False
            
    End Select
    
        Unload Me
ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_CMD:
    PopupMessage 2, err.Description, , , "Database Error"
    Resume ERR_END:
    
End Sub

Private Sub Command1_Click()

    Dim szFilename As String
    szFilename = DialogFile(Me.hWnd, 1, "Open", m_name, "Database Access" & Chr(0) & "*.mdb" & Chr(0) & "Tutti i files" & Chr(0) & "*.*", "", "mdb")
    If szFilename = "" Then Exit Sub
    
    txt_ARCHIVIO = szFilename
    
    
    DoEvents
    
    If CheckFilePath(szFilename, m_path, m_name, m_old_name) Then
            CMD_BUTTON(1).Enabled = True
        Else
            CMD_BUTTON(1).Enabled = False
    End If
   
    m_path = m_path & "\"
    txt_ARCHIVIO = m_path & m_name
    
    
    
   
End Sub



