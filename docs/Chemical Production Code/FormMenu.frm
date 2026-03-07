VERSION 5.00
Begin VB.Form FormMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AddScroll Examples"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
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
   ScaleHeight     =   2955
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLaunchExample 
      Caption         =   "Example 4"
      Height          =   645
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2145
      Width           =   3240
   End
   Begin VB.CommandButton cmdLaunchExample 
      Caption         =   "Example 3"
      Height          =   645
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1470
      Width           =   3240
   End
   Begin VB.CommandButton cmdLaunchExample 
      Caption         =   "Example 2"
      Height          =   645
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   795
      Width           =   3240
   End
   Begin VB.CommandButton cmdLaunchExample 
      Caption         =   "Example 1"
      Height          =   645
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3240
   End
End
Attribute VB_Name = "FormMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLaunchExample_Click(Index As Integer)
    Select Case Index
    Case 0
        FormExample1.Show
    Case 1
        FormExample2.DoShow
    Case 2
        FormExample3.Show
    Case 3
        Formulation.DoShow
    End Select
End Sub
