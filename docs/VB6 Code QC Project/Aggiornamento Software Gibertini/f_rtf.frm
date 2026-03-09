VERSION 5.00
Begin VB.Form f_rtf 
   Caption         =   "Specifiche Aggiornamento"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5625
   LinkTopic       =   "Form2"
   ScaleHeight     =   3705
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.OLE OLE1 
      Class           =   "AcroExch.Document.7"
      Height          =   975
      Left            =   600
      OleObjectBlob   =   "f_rtf.frx":0000
      SourceDoc       =   "C:\Documents and Settings\Utente1\Desktop\spec.pdf"
      TabIndex        =   1
      Top             =   2520
      Width           =   2895
   End
   Begin VB.OLE txAggiornameto 
      Class           =   "AcroExch.Document.7"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "f_rtf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
txAggiornameto.SourceDoc = App.Path & "\update\spec.pdf"
End Sub

Public Function DoShow() As Boolean
    Me.Visible = True
End Function

Private Sub Form_Resize()
txAggiornameto.Move 60, 60, Me.ScaleWidth - 120, Me.ScaleHeight - 120
End Sub
