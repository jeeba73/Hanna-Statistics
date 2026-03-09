VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cShowModal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const WM_DESTROY = &H2&
Private Const WM_SHOWWINDOW = &H18

Implements ISubclass

Private m_bIsModal As Boolean
Private m_hWndModal As Long
Private m_fOwner As Object

Public Sub ShowModal(ByVal fToShow As Form, fOwner As Object)
Dim hwnd As Long
   m_bIsModal = True
   fToShow.Show vbModeless, fOwner
   m_hWndModal = fToShow.hwnd
   AttachMessage Me, m_hWndModal, WM_SHOWWINDOW
   AttachMessage Me, m_hWndModal, WM_DESTROY
   Set m_fOwner = fOwner
   m_fOwner.Enabled = False
   Do While m_bIsModal
      DoEvents
   Loop
   m_fOwner.Enabled = True
   m_fOwner.ZOrder
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   ISubclass_MsgResponse = emrPostProcess
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Debug.Print "Enabling ", m_fOwner.Name, m_fOwner.hwnd, iMsg
   m_fOwner.Enabled = True
   m_fOwner.ZOrder
   DetachMessage Me, m_hWndModal, WM_SHOWWINDOW
   DetachMessage Me, m_hWndModal, WM_DESTROY
   m_bIsModal = False
End Function
