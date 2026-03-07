VERSION 5.00
Begin VB.UserControl ucScrollAdd 
   BackColor       =   &H00C0FFFF&
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   255
   ScaleWidth      =   615
   ToolboxBitmap   =   "ucScrollAdd.ctx":0000
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ScrollAdd"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   585
   End
End
Attribute VB_Name = "ucScrollAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' User control:  ucAddScroll.ctl
' Author:        Shagratt
' Dependencies:  ucScrollbar.ctl
' Last revision: 10/12/2019
' Version:       1.0.0
'========================================================================================
'
' *v1.0.0 11/12/19
' - Starting version. Basic usage working
'
'----------------------------------------------------------------------------------------
'-What it does?: It create a scrollable area on the TARGET (Control/Form) with all it's
' content .
' It add 4 controls to the Form: (exposed in ucScrollAdd.xxxx)
'   1) .PBContainer: The new container that will scroll inside the Target with all the
'      the content.
'   2) .UCScrollV: Vertical scrollbar
'   3) .UCScrollH: Horizontal scrollbar
'   4) .PBCover: A simple picture box put in bottom right corner between the scrollbars
'
'-What it dosnt?: It's not a scrollable container by itself. You need to create the
' content on a standard vb6 container (for example a PictureBox) or Form, wich has content
' larger than its viewable area. This means you can use an already made Form and add
' scrollable resize with just this control and 1 line of code.
'
'-Usage:
'   1) Add a instance on the form.
'   2) In your code call: 'ucScrollAdd.AddScroll <TARGET CONTAINER>' to automatilly add
'   scrollbars when needed. (Example: ucScrollAdd.AddScroll Me)
'
'   Optional:
'   -call 'ucScrollAdd.TrackMouseWheel' to enable MouseWheel tracking over container area.
'    (Press SHIFT to scroll horizontal)
'   -call 'ucScrollAdd.ResizeWindowLimit <MINWIDTH>, <MINHEIGHT>,[MAXWIDTH], [MAXHEIGHT]'
'    to add Form resize limit without flickering
'   -call 'ucScrollAdd.RemoveFromContainer <CTRL>,[CTRL],[CTRL],...' to exclude all
'    controls/containers you want from Scrolling.
'   -If you want to know the size of content inside the container
'    'CalcContainerContentSize maxBottom&, maxRight&' values are returned by ref.
'   -propertys ContainerW/ContainerH can be used to force the size of the container
'
'
'


Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Target As Object, TargetForm As Form
Private ControlName$, PBContainerName$, UCScrollVName$, UCScrollHName$, PBCoverName$
Public WithEvents PBContainer As PictureBox
Attribute PBContainer.VB_VarHelpID = -1
Public WithEvents UCScrollV As ucScrollbar
Attribute UCScrollV.VB_VarHelpID = -1
Public WithEvents UCScrollH As ucScrollbar
Attribute UCScrollH.VB_VarHelpID = -1
Public WithEvents PBCover As PictureBox
Attribute PBCover.VB_VarHelpID = -1
' Variable to hold 'ContainerMinW' property value
Private m_LonContainerMinW As Long
' Variable to hold 'ContainerMinH' property value
Private m_LonContainerMinH As Long
' Variable to hold 'ContainerMaxH' property value
Private m_LonContainerMaxH As Long
' Variable to hold 'ContainerMaxW' property value
Private m_LonContainerMaxW As Long

Private m_LimitResizeH As Long
Private m_LimitResizeW As Long
Private m_LimitMaxResizeH As Long
Private m_LimitMaxResizeW As Long

Private m_ResizePaddingBottom As Long 'Adjust BotRight when resizing
Private m_ResizePaddingRight As Long
Private m_OriginalTargetW As Long 'Original Target Size
Private m_OriginalTargetH As Long


'Internal use
Private SubClassStarted As Boolean 'Flag
Private ScrollbarAdjusted As Boolean 'Flag

Public Enum eVorH
    None = 0
    Vertical = 1
    Horizontal = 2
End Enum

Public Enum eVorHB
    Both = 0
    Only_Vertical = 1
    Only_Horizontal = 2
End Enum


Private TrackingDir As eVorH
Private m_LonScrollBarSizePX As Long

'-- Events:
Public Event ScrollV(Value&)
Public Event ScrollH(Value&)


'-Selfsub declarations----------------------------------------------------------------------------
Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
End Enum

Private Type tSubData                                                         'Subclass data type
    hWnd                   As Long                                            'Handle of the window being subclassed
    nAddrSub               As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig              As Long                                            'The address of the pre-existing WndProc
    nMsgCntA               As Long                                            'Msg after table entry count
    nMsgCntB               As Long                                            'Msg before table entry count
    aMsgTblA()             As Long                                            'Msg after table array
    aMsgTblB()             As Long                                            'Msg Before table array
End Type

Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback
Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries
Private Const WNDPROC_OFF   As Long = &H38                                  'Thunk offset to the WndProc execution address
Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN  As Long = 1                                     'Thunk data index of the shutdown flag
Private Const IDX_HWND      As Long = 2                                     'Thunk data index of the subclassed hWnd
Private Const IDX_WNDPROC   As Long = 9                                     'Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11                                    'Thunk data index of the Before table
Private Const IDX_ATABLE    As Long = 12                                    'Thunk data index of the After table
Private Const IDX_PARM_USER As Long = 13                                    'Thunk data index of the User-defined callback parameter data index

Private z_ScMem             As Long                                         'Thunk base address
Private z_Sc(64)            As Long                                         'Thunk machine-code initialised here
Private z_Funk              As Collection                                   'hWnd/thunk-address collection

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'*2-
'To limit resize
Private Const WM_SIZE As Long = &H5
Private Const WM_SYSCOMMAND As Long = &H112
Private Const SC_MAXIMIZE As Long = &HF030&
Private Const WM_GETMINMAXINFO As Long = &H24
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MINMAXINFO
  ptReserved As POINTAPI
  ptMaxSize As POINTAPI
  ptMaxPosition As POINTAPI
  ptMinTrackSize As POINTAPI
  ptMaxTrackSize As POINTAPI
End Type
'To limit resize
' Variable to hold 'AutoScrollbars' property value
Private m_AutoScrollbars As eVorHB

Public Property Get AutoScrollbars() As eVorHB
    AutoScrollbars = m_AutoScrollbars
End Property
Public Property Let AutoScrollbars(ByVal EVoValue As eVorHB)
    m_AutoScrollbars = EVoValue
    PropertyChanged "AutoScrollbars"
End Property

'Give the scrollbar size on form scale (we asume its simetric)
'// Runtime read only
Public Property Get ScrollBarSize() As Long
    ScrollBarSize = SCX(m_LonScrollBarSizePX)
End Property

Public Property Get ScrollBarSizePX() As Long
    ScrollBarSizePX = m_LonScrollBarSizePX
End Property
Public Property Let ScrollBarSizePX(ByVal LonValue As Long)
    m_LonScrollBarSizePX = LonValue
    PropertyChanged "ScrollBarSizePX"
End Property

'// Runtime read only
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'Limit resize of the container [Width]
Public Property Get ContainerW() As Long
On Error Resume Next
    If (PBContainer Is Nothing) Then Exit Property
    ContainerW = PBContainer.Width
End Property

Public Property Let ContainerW(ByVal LonValue As Long)
On Error GoTo err:
Dim OldValue&
    If (PBContainer Is Nothing) Then Exit Property
    If (LonValue < m_LonContainerMinW) And (m_LonContainerMinW <> -1) Then LonValue = m_LonContainerMinW
    If (LonValue > m_LonContainerMaxW) And (m_LonContainerMaxW <> -1) Then LonValue = m_LonContainerMaxW
    OldValue& = PBContainer.Width
    If (OldValue <> LonValue) Then
        PBContainer.Width = LonValue
        AdjustScrollBars
    End If
err:
End Property

Public Property Get ContainerMaxW() As Long
    ContainerMaxW = m_LonContainerMaxW
End Property
Public Property Let ContainerMaxW(ByVal LonValue As Long)
    m_LonContainerMaxW = LonValue
    If (ContainerW > LonValue) Then ContainerW = LonValue
    PropertyChanged "ContainerMaxW"
End Property

Public Property Get ContainerMinW() As Long
    ContainerMinW = m_LonContainerMinW
End Property
Public Property Let ContainerMinW(ByVal LonValue As Long)
    m_LonContainerMinW = LonValue
    If (ContainerW < LonValue) Then ContainerW = LonValue
    PropertyChanged "ContainerMinW"
End Property



'Limit resize of the container [Height]
Public Property Get ContainerH() As Long
On Error Resume Next
    ContainerH = PBContainer.Height
End Property
Public Property Let ContainerH(ByVal LonValue As Long)
On Error GoTo err:
Dim OldValue&
    If (PBContainer Is Nothing) Then Exit Property
    If (LonValue < m_LonContainerMinH) And (m_LonContainerMinH <> -1) Then LonValue = m_LonContainerMinH
    If (LonValue > m_LonContainerMaxH) And (m_LonContainerMaxH <> -1) Then LonValue = m_LonContainerMaxH
    OldValue& = PBContainer.Height
    If (OldValue <> LonValue) Then
        PBContainer.Height = LonValue
        AdjustScrollBars
    End If
err:
End Property

Public Property Get ContainerMaxH() As Long
    ContainerMaxH = m_LonContainerMaxH
End Property
Public Property Let ContainerMaxH(ByVal LonValue As Long)
    m_LonContainerMaxH = LonValue
    If (ContainerH > LonValue) Then ContainerH = LonValue
    PropertyChanged "ContainerMaxH"
End Property

Public Property Get ContainerMinH() As Long
    ContainerMinH = m_LonContainerMinH
End Property
Public Property Let ContainerMinH(ByVal LonValue As Long)
    m_LonContainerMinH = LonValue
    If (ContainerH < LonValue) Then ContainerH = LonValue
    PropertyChanged "ContainerMinH"
End Property


'=======================================================================================
'Enable scrolling while the mouse is over the Target area (not only over the scrollbars)
'=======================================================================================
Public Sub TrackMouseWheel(Optional lTrackingDir As eVorH = eVorH.Vertical)
On Error GoTo err:
    If (lTrackingDir = Vertical) Then
        UCScrollV.TrackMouseWheelOnHwnd Target.hWnd ' UCScrollV.Parent.hwnd ' PBContainer.hwnd
    ElseIf (lTrackingDir = Horizontal) Then
        UCScrollH.TrackMouseWheelOnHwnd UCScrollH.Parent.hWnd 'PBContainer.hwnd
    Else
        'Disable tracking
        If (TrackingDir = Vertical) Then
            UCScrollV.TrackMouseWheelOnHwndStop 'PBContainer.hwndPBContainer.hwnd
        ElseIf (TrackingDir = Horizontal) Then
            UCScrollH.TrackMouseWheelOnHwndStop 'PBContainer.hwndPBContainer.hwnd
        Else
            'Was not doing any ext tracking so dont do anything
        End If
    End If
    TrackingDir = lTrackingDir
    Exit Sub
err:
    Debug.Print ("TrackMouseWheel Err: " & err.Description)
End Sub





'==============================================
' Add the scrollbars to the target Control/Form
'==============================================
Public Sub AddScroll(lTarget As Object)
On Error GoTo err
Dim ctl As Control
Dim bPBAlreadyExist
Dim calcRight&, maxRight&, calcBottom&, maxBottom&

    'TODO: verificar que tenga hWnd
    On Error Resume Next
    err.Clear
    Dim l&
    l = lTarget.hWnd
    If (err.NUMBER > 0) Then
        Debug.Print (UserControl.Extender.Name & " Error: Target not have hWnd. Aborting")
    End If
    On Error GoTo err
    
    
    'Create the names of the controls I will add (a picturebox as a container and 2 scrollbars)
    ControlName$ = UserControl.Ambient.DisplayName
    PBContainerName$ = ControlName$ & "_pbScrollContainer"
    UCScrollVName$ = ControlName$ & "_ScrollV"
    UCScrollHName$ = ControlName$ & "_ScrollH"
    PBCoverName$ = ControlName$ & "_pbCover"
    
    Set Target = lTarget 'target object (Form/Picbox/etc)
    Set TargetForm = pFindFormObj(Target) ' owner form of Target
    
    '1) Sanity check. Container and scrollbars already added?
    For Each ctl In TargetForm.Controls
        If (ctl.Name = PBContainerName$) Then
            Debug.Print (UserControl.Extender.Name & " Error: Container '" & PBContainerName$ & "' already present. Aborting")
            AdjustScrollBars
            bPBAlreadyExist = True
            Exit Sub
        End If
    Next ctl
    
    '2) Create a new Picturebox to be used as container and place it inside the target
    Set PBContainer = TargetForm.Controls.Add("vb.PictureBox", PBContainerName$)
    SetParent PBContainer.hWnd, Target.hWnd
    
    '3) Check maximum Bottom and Right of the controls inside the target (Asume all start from 0,0)
    '4) and also change its container to the new Container (Picturebox) we created.
    On Error Resume Next
    CalcContainerContentSize maxBottom&, maxRight&, True
    
    
    '5) With everything calculated we set the properties and size of our container
    On Error Resume Next
    PBContainer.Appearance = 0 'Flat
    PBContainer.BackColor = Target.BackColor
    Set PBContainer.Picture = Target.Picture
    'PBContainer.BorderStyle = BorderStyleConstants.vbBSSolid
    PBContainer.BorderStyle = BorderStyleConstants.vbTransparent
    PBContainer.Top = 0
    PBContainer.Left = 0
    PBContainer.Height = maxBottom
    PBContainer.Width = maxRight
    PBContainer.AutoRedraw = False
    PBContainer.Visible = True
    err.Clear

    '6) Add Scrollbars
    'Vertical
    Set UCScrollV = TargetForm.Controls.Add("ChemicalProduction.ucScrollBar", UCScrollVName$)
    SetParent UCScrollV.hWnd, Target.hWnd
    UCScrollV.AddedFromCodeINIT_AND_DEFAULTS
    UCScrollV.Style = sFlat
    If (err.NUMBER <> 0) Then Debug.Print ("ChemicalProduction UCScrollV Error: " & err.Description) 'Capture error on create
    
    'Horizontal
    Set UCScrollH = TargetForm.Controls.Add("ChemicalProduction.ucScrollBar", UCScrollHName$)
    SetParent UCScrollH.hWnd, Target.hWnd
    UCScrollH.AddedFromCodeINIT_AND_DEFAULTS
    UCScrollH.Style = sFlat
    UCScrollH.Orientation = oHorizontal
    UCScrollH.DisableMouseWheelSupport = True
    If (err.NUMBER <> 0) Then Debug.Print ("ChemicalProduction UCScrollH Error: " & err.Description) 'Capture error on prop.set
    err.Clear
    
    'Link both scrollbars (for transfering scroll on SHIFT pressed)
    UCScrollV.AttachHorizontalScrollBar UCScrollH
    
    
    '7) If the user dont set any limit then set the minimun size of container to not show the scrollbars
    'Reminder: This is the container, not the visible area wich default to the target
    '          area and its always managed by the user form. (with the exception when
    '          ResizeTargetOnFormResize is used or Target=Form)
    If (ContainerMinH = -1) Then ContainerMinH = maxBottom
    If (ContainerMinW = -1) Then ContainerMinW = maxRight
    
    '8) Create another Picturebox to cover the gap between the 2 scrollbars
    Set PBCover = TargetForm.Controls.Add("vb.PictureBox", PBCoverName$)
    SetParent PBCover.hWnd, Target.hWnd
    
    PBCover.BorderStyle = 0
    
    '9) Put the scrollbars on top
    UCScrollV.ZOrder
    UCScrollH.ZOrder
    PBCover.ZOrder
    
    '10) Finally check positions and if its needed to show the scrollbars/cover
    AdjustScrollBars
    
    'If we detect we are adding scrollbars to the form itself, we active the
    'auto reposition of the scrollbars when the form is resized
    If (Target Is TargetForm) Then
        ResizeTargetOnFormResize 0, 0
    End If
    
    If (err.NUMBER <> 0) Then Debug.Print ("AddScroll end with Error: " & err.Description)
    Exit Sub
err:
    Debug.Print ("AddScroll Error: " & err.Description)
End Sub

'==========================================================
'Check if the scrollbars need to be show and adjust its
'position and size
'==========================================================
'It's public so it can be called by form if needed
Public Sub AdjustScrollBars()
On Error GoTo err:
Dim auxV&, auxH&, aux&, auxd As Double

    'Sanity check
    If (UCScrollV Is Nothing) Then Exit Sub
    If (TargetForm.WindowState = vbMinimized) Then Exit Sub
    If (Target.Width < SCX(25)) Or (Target.Height < SCY(45)) Then Exit Sub
    
    ScrollbarAdjusted = True
    'Check if after resizing form we need to move form down
    '(case when we scrolled down but now after resize we got space for everything to be displayed on screen)
    If (PBContainer.Top < 0) Then
        aux& = (Target.ScaleHeight + Abs(PBContainer.Top)) - PBContainer.Height
        If (aux > 0) Then
            PBContainer.Top = PBContainer.Top + aux
        End If
    End If
        
    
    'Padding bot/right disabled?
    If (m_ResizePaddingBottom <= 0) And (m_ResizePaddingRight <= 0) Then
        UCScrollH.Extender.Height = ScrollBarSize
        UCScrollH.Top = Target.ScaleHeight - ScrollBarSize
        UCScrollH.Left = 0
        UCScrollH.Extender.Width = Target.ScaleWidth - ScrollBarSize
    
        UCScrollV.Top = 0
        UCScrollV.Left = Target.ScaleWidth - ScrollBarSize
        UCScrollV.Extender.Height = Target.ScaleHeight - ScrollBarSize
        UCScrollV.Extender.Width = ScrollBarSize
    Else
        'Apply padding
        UCScrollH.Extender.Height = ScrollBarSize
        UCScrollH.Top = Target.ScaleHeight - (ScrollBarSize + SCY(m_ResizePaddingBottom))
        UCScrollH.Left = 0
        UCScrollH.Extender.Width = Target.ScaleWidth - (ScrollBarSize + SCX(m_ResizePaddingRight))
    
        UCScrollV.Top = 0
        UCScrollV.Left = Target.ScaleWidth - (ScrollBarSize + SCX(m_ResizePaddingRight))
        UCScrollV.Extender.Height = Target.ScaleHeight - (ScrollBarSize + SCY(m_ResizePaddingBottom))
        UCScrollV.Extender.Width = ScrollBarSize
    End If
    
    'Need Vertical Scroll ?
    auxV& = PBContainer.Height - Target.ScaleHeight
    If (auxV > 0) And Not (m_AutoScrollbars = eVorHB.Only_Horizontal) Then
        UCScrollV.Min = 0
        UCScrollV.Max = auxV
        aux = CDbl(auxV * (Target.ScaleHeight / PBContainer.Height))
        If (aux = 0) Then aux = 1
        UCScrollV.LargeChange = aux
        UCScrollV.SmallChange = UCScrollV.LargeChange / 10
        UCScrollV.Visible = True
    Else
        auxV = 0
        UCScrollV.Visible = False
    End If

    'Need Horizontal Scroll ?
    'Take xtra space for the vertical bar
    If (auxV > 0) Then
        auxH& = (PBContainer.Width + ScrollBarSize) - Target.ScaleWidth
    Else
        auxH& = PBContainer.Width - Target.ScaleWidth
    End If
    
    If (auxH > 0) And Not (m_AutoScrollbars = eVorHB.Only_Vertical) Then
        UCScrollH.Min = 0
        UCScrollH.Max = auxH
        'UCScrollH.LargeChange = auxH * (Target.ScaleWidth / PBContainer.Width)
        aux = CDbl(auxH * (Target.ScaleWidth / PBContainer.Width))
        If (aux = 0) Then aux = 1
        UCScrollH.LargeChange = aux
        UCScrollH.SmallChange = UCScrollH.LargeChange / 10
        UCScrollH.Visible = True
    Else
        auxH = 0
        UCScrollH.Visible = False
    End If
    
    'Recheck if after adding Horizontal scroll we need a vertical one
    If (auxV < 0) And (auxH > 0) Then
        auxV& = (PBContainer.Height + ScrollBarSize + SCY(5)) - Target.ScaleHeight
        If (auxV > 0) Then
            UCScrollV.Min = 0
            UCScrollV.Max = auxV
            UCScrollV.LargeChange = auxV * (Target.ScaleHeight / PBContainer.Height)
            UCScrollV.SmallChange = UCScrollV.LargeChange / 10
            UCScrollV.Visible = True
        End If
    End If

    'If there is only one ScrollBar extend its size
    'Extend Vertical
    If (auxV > 0) And (auxH <= 0) Then
        UCScrollV.Extender.Height = Target.ScaleHeight
    'Extend Horizontal
    ElseIf (auxH > 0) And (auxV <= 0) Then
        UCScrollH.Extender.Width = Target.ScaleWidth
    End If
    
    '2 Scrollbars -> Display cover in bottom/right
    If (auxH > 1) And (auxV > 1) Then
        PBCover.Top = UCScrollH.Top
        PBCover.Left = UCScrollV.Left
        If (PBCover.Width <> ScrollBarSize) Then
            PBCover.Width = ScrollBarSize
            PBCover.Height = ScrollBarSize
        End If
        PBCover.Visible = True
    Else
        PBCover.Visible = False
    End If

    Exit Sub
err:
    Debug.Print (UserControl.Extender.Name & ".AdjustScrollBars Err: " & err.Description)
End Sub

'=============================================================
'Calculate the size of the content inside the target
'=============================================================
'Note: Return values by ref
Public Sub CalcContainerContentSize(ByRef maxBottom&, ByRef maxRight&, Optional bChangeContainer As Boolean = False)
Dim ctl As Control
Dim calcRight&, calcBottom&
Dim oh&

    'Check Bottom and Right of every control inside the target (Asume all start from 0,0)
    On Error Resume Next
    For Each ctl In TargetForm.Controls
    
        err.Clear
        oh = 0&
        
        'All exceptions should be added
        Select Case TypeName(ctl)
        Case "Timer"
            'Dont process
        Case Else
            oh = GetParent(ctl.hWnd)
            If (err.NUMBER <> 0) Then
                'Control dont support hwnd? (Ex: labels) let's try getting parent from its container
                err.Clear
                oh = ctl.Container.hWnd
                If (err.NUMBER <> 0) Then
                    Debug.Print ("CalcContainerContentSize Err: Fail to get parent from " & ctl.Name)
                    err.Clear
                End If
            End If
        End Select
        
        If Not ((ctl Is PBContainer) Or (ctl Is Me) Or (ctl Is Target) Or (oh = 0)) Then
            If (oh = Target.hWnd) Then
                'sLog "AddScroll changing container of " & ctl.Name & " to " & Target.Name
                
                ''DEBUG
                'If (ctl.Name = "TimMovTitle1") Then
                '    ctl.Name = ctl.Name
                'End If
                
                'Change the container (should be used only from AddScroll)
                If (bChangeContainer) Then Set ctl.Container = PBContainer
                
                'Controls with diferent calculation of size should be added
                Select Case (TypeName(ctl))
                Case "Line"
                    calcBottom = ctl.y2
                    calcRight = ctl.x2
                Case Else
                    calcBottom = ctl.Top + ctl.Height + SCY(15)
                    calcRight = ctl.Left + ctl.Width
                End Select
                
                'Keep the max values
                If (calcBottom > maxBottom) Then maxBottom = calcBottom
                If (calcRight > maxRight) Then maxRight = calcRight
            
                If (err.NUMBER <> 0) Then
                    Debug.Print ("CalcContainerContentSize ctl: " & ctl.Name & ", Error: " & err.Description)
                    err.Clear
                End If
            End If
        End If

    Next ctl
End Sub

Private Sub UCScrollV_Change()
    'sLog UCScrollV.Value
    PBContainer.Top = -UCScrollV.Value
    RaiseEvent ScrollV(UCScrollV.Value)
    
End Sub

Private Sub UCScrollH_Change()
    'sLog UCScrollH.Value
    PBContainer.Left = -UCScrollH.Value
    RaiseEvent ScrollH(UCScrollH.Value)
End Sub

'*2-
'===================================
'Enable TargetForm minimal Resize Limit
'NOTE: set -1,-1 to disable
'===================================
'Size needs to include window borders and title!!!
Public Function ResizeWindowLimit(Optional MinWidthPX& = -1, Optional MinHeightPX& = -1, Optional MaxWidthPX& = -1, Optional MaxHeightPX& = -1)
    m_LimitResizeH = MinHeightPX&
    m_LimitResizeW = MinWidthPX&
    m_LimitMaxResizeH = MaxHeightPX&
    m_LimitMaxResizeW = MaxWidthPX&
    If Not (SubClassStarted) Then StartSubclassing
End Function

Public Function Terminate()
UserControl_Terminate

End Function

'===================================
'Enable Reposition of scrollbars
'NOTE: set -1,-1 to disable
'===================================
Public Function ResizeTargetOnFormResize(Optional PaddingBotPX& = 0, Optional PaddingRightPX& = 0)
    m_ResizePaddingBottom = PaddingBotPX&
    m_ResizePaddingRight = PaddingRightPX&
    If Not (SubClassStarted) Then StartSubclassing
End Function


'===============================================
'Remove items from the Scrollbar container
'so they dont scroll
'===============================================
'Example:
'ucScrollAdd1.RemoveFromContainer PBTitle, PBHover
Public Sub RemoveFromContainer(ParamArray Param())
Dim o As Variant
On Error Resume Next

    For Each o In Param
        Set o.Container = TargetForm
        o.ZOrder
    Next
    
    'Put the scrollbar on top again
    UCScrollH.ZOrder
    UCScrollV.ZOrder
    
End Sub

'=========================================================
'Return the form containing the control
'=========================================================
Public Function pFindFormObj(c As Object) As Form
On Error GoTo Err1
Dim oContenedor As Object

    'If passed object is already a form return it
    If (TypeOf c Is Form) Then
        Set pFindFormObj = c
        Exit Function
    End If
    Set oContenedor = c.Container

    'Loop all containers
    Do
        Set oContenedor = oContenedor.Container
    Loop

Err1:
    If (TypeOf oContenedor Is Form) Then
        Set pFindFormObj = oContenedor
        Exit Function
    End If
End Function

'===============================================
'Scale X,Y from pixels in Target Form scale
'===============================================
Private Function SCX(Value&) As Long
    If (TargetForm Is Nothing) Then
        Dim f As Form
        Set f = pFindFormObj(UserControl.Parent)
        SCX = f.ScaleX(Value, ScaleModeConstants.vbPixels, f.ScaleMode)
    Else
        SCX = TargetForm.ScaleX(Value, ScaleModeConstants.vbPixels, TargetForm.ScaleMode)
    End If
End Function
Private Function SCY(Value&) As Long
    If (TargetForm Is Nothing) Then
        Dim f As Form
        Set f = pFindFormObj(UserControl.Parent)
        SCY = f.ScaleY(Value, ScaleModeConstants.vbPixels, f.ScaleMode)
    Else
        SCY = TargetForm.ScaleY(Value, ScaleModeConstants.vbPixels, TargetForm.ScaleMode)
    End If
End Function


Private Sub UserControl_Resize()
    UserControl.Width = 645
    UserControl.Height = 240
End Sub


Private Sub UserControl_Terminate()
On Error Resume Next

    TrackMouseWheel None
    If Not (UCScrollV Is Nothing) Then Set UCScrollV = Nothing
    If Not (UCScrollH Is Nothing) Then Set UCScrollH = Nothing
    If Not (PBCover Is Nothing) Then Set PBCover = Nothing
    If Not (PBContainer Is Nothing) Then Set PBContainer = Nothing
    DoEvents
    StopSubclassing
    DoEvents
End Sub


Private Sub UserControl_Initialize()
    m_LimitResizeH = -1
    m_LimitResizeW = -1
    m_LimitMaxResizeH = -1
    m_LimitMaxResizeW = -1
    m_ResizePaddingBottom = -1
    m_ResizePaddingRight = -1
End Sub

Private Sub UserControl_InitProperties()
    m_LonContainerMinW = -1
    m_LonContainerMinH = -1
    m_LonContainerMaxH = -1
    m_LonContainerMaxW = -1
    m_LonScrollBarSizePX = 15
    m_AutoScrollbars = eVorHB.Both
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_LonContainerMinW = PropBag.ReadProperty("ContainerMinW", -1)
    m_LonContainerMinH = PropBag.ReadProperty("ContainerMinH", -1)
    m_LonContainerMaxH = PropBag.ReadProperty("ContainerMaxH", -1)
    m_LonContainerMaxW = PropBag.ReadProperty("ContainerMaxW", -1)
    m_LonScrollBarSizePX = PropBag.ReadProperty("ScrollBarSizePX", 15)
    
    'NOTE: Commented cause we need to wait for TargetForm to be set
    ''-- Run-time?
    'If (Ambient.UserMode) Then
    '    Call StartSubclassing
    'End If
    
    m_AutoScrollbars = PropBag.ReadProperty("AutoScrollbars", eVorHB.Both)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ContainerMinW", m_LonContainerMinW, -1)
    Call PropBag.WriteProperty("ContainerMinH", m_LonContainerMinH, -1)
    Call PropBag.WriteProperty("ContainerMaxH", m_LonContainerMaxH, -1)
    Call PropBag.WriteProperty("ContainerMaxW", m_LonContainerMaxW, -1)
    Call PropBag.WriteProperty("ScrollBarSizePX", m_LonScrollBarSizePX, 15)
    Call PropBag.WriteProperty("AutoScrollbars", m_AutoScrollbars, eVorHB.Both)
End Sub


'Replace with prefered subclassing
Public Sub StartSubclassing()
    
    If (SubClassStarted) Then Exit Sub
    If (TargetForm Is Nothing) Then Exit Sub
    SubClassStarted = True
    
    'Paul Caton SC
    Call sc_Subclass(TargetForm.hWnd) 'Optional pass and object with ObjPtr(pic(0))
    Call sc_AddMsg(TargetForm.hWnd, WM_GETMINMAXINFO, [MSG_BEFORE])
    Call sc_AddMsg(TargetForm.hWnd, WM_SIZE, [MSG_AFTER])
    Call sc_AddMsg(TargetForm.hWnd, WM_SYSCOMMAND, [MSG_AFTER])
    
End Sub

Public Sub StopSubclassing()
    If (SubClassStarted) Then
        'Paul Caton SC
        sc_Terminate
    End If
End Sub

'*-
'-SelfSub code------------------------------------------------------------------------------------
'==================================================================================================
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
'* v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
'* v1.1 VirtualAlloc memory to prevent Data Execution Prevention faults on Win64......... 20060324
'* v1.2 Thunk redesigned to handle unsubclassing and memory release...................... 20060325
'* v1.3 Data array scrapped in favour of property accessors.............................. 20060405
'* v1.4 Optional IDE protection added
'*      User-defined callback parameter added
'*      All user routines that pass in a hWnd get additional validation
'*      End removed from zError.......................................................... 20060411
'* v1.5 Added nOrdinal parameter to sc_Subclass
'*      Switched machine-code array from Currency to Long................................ 20060412
'* v1.6 Added an optional callback target object
'*      Added an IsBadCodePtr on the callback address in the thunk prior to callback..... 20060413
'*************************************************************************************************
Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True) As Boolean 'Subclass the specified window handle
'*************************************************************************************************
'* lng_hWnd   - Handle of the window to subclass
'* lParamUser - Optional, user-defined callback parameter
'* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
'*************************************************************************************************
Const CODE_LEN      As Long = 260                                           'Thunk length in bytes
Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))            'Bytes to allocate per thunk, data + code + msg tables
Const PAGE_RWX      As Long = &H40&                                         'Allocate executable memory
Const MEM_COMMIT    As Long = &H1000&                                       'Commit allocated memory
Const MEM_RELEASE   As Long = &H8000&                                       'Release allocated memory flag
Const IDX_EBMODE    As Long = 3                                             'Thunk data index of the EbMode function address
Const IDX_CWP       As Long = 4                                             'Thunk data index of the CallWindowProc function address
Const IDX_SWL       As Long = 5                                             'Thunk data index of the SetWindowsLong function address
Const IDX_FREE      As Long = 6                                             'Thunk data index of the VirtualFree function address
Const IDX_BADPTR    As Long = 7                                             'Thunk data index of the IsBadCodePtr function address
Const IDX_OWNER     As Long = 8                                             'Thunk data index of the Owner object's vTable address
Const IDX_CALLBACK  As Long = 10                                            'Thunk data index of the callback method address
Const IDX_EBX       As Long = 16                                            'Thunk code patch index of the thunk data
Const SUB_NAME      As String = "sc_Subclass"                               'This routine's name
  Dim nAddr         As Long
  Dim nID           As Long
  Dim nMyID         As Long
  
  If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
    zError SUB_NAME, "Invalid window handle"
    Exit Function
  End If

  nMyID = GetCurrentProcessId                                               'Get this process's ID
  GetWindowThreadProcessId lng_hWnd, nID                                    'Get the process ID associated with the window handle
  If nID <> nMyID Then                                                      'Ensure that the window handle doesn't belong to another process
    zError SUB_NAME, "Window handle belongs to another process"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the address of the specified ordinal method
  If nAddr = 0 Then                                                         'Ensure that we've found the ordinal method
    zError SUB_NAME, "Callback method not found"
    Exit Function
  End If
    
  If z_Funk Is Nothing Then                                                 'If this is the first time through, do the one-time initialization
    Set z_Funk = New Collection                                             'Create the hWnd/thunk-address collection
    z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
    z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&

    z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                    'Store CallWindowProc function address in the thunk data
    z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                     'Store the SetWindowLong function address in the thunk data
    z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                     'Store the VirtualFree function address in the thunk data
    z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")                  'Store the IsBadCodePtr function address in the thunk data
  End If
  
  z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                  'Allocate executable memory

  If z_ScMem <> 0 Then                                                      'Ensure the allocation succeeded
    On Error GoTo CatchDoubleSub                                            'Catch double subclassing
      z_Funk.Add z_ScMem, "h" & lng_hWnd                                    'Add the hWnd/thunk-address to the collection
    On Error GoTo 0
  
    If bIdeSafety Then                                                      'If the user wants IDE protection
      z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                          'Store the EbMode function address in the thunk data
    End If
    
    z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
    z_Sc(IDX_HWND) = lng_hWnd                                               'Store the window handle in the thunk data
    z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
    z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
    z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
    z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
    z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
    
    nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
    If nAddr = 0 Then                                                       'Ensure the new WndProc was set correctly
      zError SUB_NAME, "SetWindowLong failed, error #" & err.LastDllError
      GoTo ReleaseMemory
    End If
        
    z_Sc(IDX_WNDPROC) = nAddr                                               'Store the original WndProc address in the thunk data
    RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
    sc_Subclass = True                                                      'Indicate success
  Else
    zError SUB_NAME, "VirtualAlloc failed, error: " & err.LastDllError
  End If
  
  Exit Function                                                             'Exit sc_Subclass

CatchDoubleSub:
  zError SUB_NAME, "Window handle is already subclassed"
  
ReleaseMemory:
  VirtualFree z_ScMem, 0, MEM_RELEASE                                       'sc_Subclass has failed after memory allocation, so release the memory
End Function

'Terminate all subclassing
Private Sub sc_Terminate()
  Dim i As Long

  If Not (z_Funk Is Nothing) Then                                           'Ensure that subclassing has been started
    With z_Funk
      For i = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order
        z_ScMem = .Item(i)                                                  'Get the thunk address
        If IsBadCodePtr(z_ScMem) = 0 Then                                   'Ensure that the thunk hasn't already released its memory
          sc_UnSubclass zData(IDX_HWND)                                     'UnSubclass
        End If
      Next i                                                                'Next member of the collection
    End With
    Set z_Funk = Nothing                                                    'Destroy the hWnd/thunk-address collection
  End If
End Sub

'UnSubclass the specified window handle
Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "sc_UnSubclass", "Window handle isn't subclassed"
  Else
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                           'Ensure that the thunk hasn't already released its memory
      zData(IDX_SHUTDOWN) = -1                                              'Set the shutdown indicator
      zDelMsg ALL_MESSAGES, IDX_BTABLE                                      'Delete all before messages
      zDelMsg ALL_MESSAGES, IDX_ATABLE                                      'Delete all after messages
    End If
    z_Funk.Remove "h" & lng_hWnd                                            'Remove the specified window handle from the collection
  End If
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be added to the before original WndProc table...
      zAddMsg uMsg, IDX_BTABLE                                              'Add the message to the before table
    End If
    If When And MSG_AFTER Then                                              'If message is to be added to the after original WndProc table...
      zAddMsg uMsg, IDX_ATABLE                                              'Add the message to the after table
    End If
  End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be deleted from the before original WndProc table...
      zDelMsg uMsg, IDX_BTABLE                                              'Delete the message from the before table
    End If
    If When And MSG_AFTER Then                                              'If the message is to be deleted from the after original WndProc table...
      zDelMsg uMsg, IDX_ATABLE                                              'Delete the message from the after table
    End If
  End If
End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_CallOrigWndProc = _
        CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
  End If
End Function

'Get the subclasser lParamUser callback parameter
Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_lParamUser = zData(IDX_PARM_USER)                                    'Get the lParamUser callback parameter
  End If
End Property

'Let the subclasser lParamUser callback parameter
Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal NewValue As Long)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    zData(IDX_PARM_USER) = NewValue                                         'Set the lParamUser callback parameter
  End If
End Property

'-The following routines are exclusively for the sc_ subclass routines----------------------------

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim i      As Long                                                        'Loop index

  nBase = z_ScMem                                                            'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                    'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
    nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
  Else
    nCount = zData(0)                                                       'Get the current table entry count
    If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
      GoTo Bail
    End If

    For i = 1 To nCount                                                     'Loop through the table entries
      If zData(i) = 0 Then                                                  'If the element is free...
        zData(i) = uMsg                                                     'Use this element
        GoTo Bail                                                           'Bail
      ElseIf zData(i) = uMsg Then                                           'If the message is already in the table...
        GoTo Bail                                                           'Bail
      End If
    Next i                                                                  'Next message table entry

    nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count
    zData(nCount) = uMsg                                                    'Store the message in the appended table entry
  End If

  zData(0) = nCount                                                         'Store the new table entry count
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim i      As Long                                                        'Loop index

  nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                   'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
    zData(0) = 0                                                            'Zero the table entry count
  Else
    nCount = zData(0)                                                       'Get the table entry count
    
    For i = 1 To nCount                                                     'Loop through the table entries
      If zData(i) = uMsg Then                                               'If the message is found...
        zData(i) = 0                                                        'Null the msg value -- also frees the element for re-use
        GoTo Bail                                                           'Bail
      End If
    Next i                                                                  'Next message table entry
    
    zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
  End If
  
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map zData() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "zMap_hWnd", "Subclassing hasn't been started"
  Else
    On Error GoTo Catch                                                     'Catch unsubclassed window handles
    z_ScMem = z_Funk("h" & lng_hWnd)                                        'Get the thunk address
    zMap_hWnd = z_ScMem
  End If
  
  Exit Function                                                             'Exit returning the thunk address

Catch:
  zError "zMap_hWnd", "Window handle isn't subclassed"
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, i, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  j = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < j
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 128                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop

End Function

Private Property Get zData(ByVal nIndex As Long) As Long
  RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
  RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property

'*-
'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lhWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Object)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************
On Error GoTo err
    
    Select Case uMsg
    Case WM_GETMINMAXINFO
    
        If (m_LimitResizeH > -1) And (m_LimitResizeW > -1) Then
            If (TargetForm.WindowState <> vbNormal) Then Exit Sub
            ' dimention a variable to hold the structure passed from Windows in lParam
            Dim udtMINMAXINFO As MINMAXINFO
            Dim nWidthPixels As Long, nHeightPixels As Long
    
            'nWidthPixels = Screen.Width * Screen.TwipsPerPixelX
            'nHeightPixels = Screen.Height * Screen.TwipsPerPixelY
    
            ' copy the struct to our UDT variable
            CopyMemory udtMINMAXINFO, ByVal lParam, 40&
    
            With udtMINMAXINFO
                ' set the width of the form when it's maximized
                '.ptMaxSize.x = 500
                ' set the height of the form when it's maximized
                '.ptMaxSize.y = 500
                'Debug.Print nWidthPixels
                'set the Left of the form when it's maximized
                '.ptMaxPosition.x = nWidthPixels * 8
                ' set the Top of the form when it's maximized
                '.ptMaxPosition.y = nHeightPixels * 8
    
                If (m_LimitMaxResizeH > -1) And (m_LimitMaxResizeW > -1) Then
                    ' set the max width that the user can size the form
                    .ptMaxTrackSize.X = m_LimitMaxResizeW
                    ' set the max height that the user can size the form
                    .ptMaxTrackSize.Y = m_LimitMaxResizeH
                End If
    
                ' set the min width that the user can size the form
                .ptMinTrackSize.X = m_LimitResizeW
                ' set the min height that the user can size the form
                .ptMinTrackSize.Y = m_LimitResizeH
                
            End With
    
            'Copy our modified struct back to the Windows struct
            CopyMemory ByVal lParam, udtMINMAXINFO, 40&
    
            'Return zero indicating that we have acted on this message
            bHandled = True: lReturn = 0
        End If
    
    Case WM_SIZE
        If (m_ResizePaddingBottom >= 0) And (m_ResizePaddingRight >= 0) Then
            If Not (Target Is TargetForm) Then
                Target.Width = TargetForm.ScaleWidth - Target.Left
                Target.Height = TargetForm.ScaleHeight - Target.Top
            End If
        End If
        
        'Set the flag so we detect if changing the container width result in calling AdjustScrollBars
        'and dont call it again.
        ScrollbarAdjusted = False
        ContainerW = TargetForm.ScaleWidth - ScrollBarSize
        If Not (ScrollbarAdjusted) Then AdjustScrollBars
        
        
        
    Case WM_SYSCOMMAND
        If (wParam = SC_MAXIMIZE) Then
            ScrollbarAdjusted = False
            ContainerW = TargetForm.ScaleWidth - ScrollBarSize
            If Not (ScrollbarAdjusted) Then AdjustScrollBars
        End If
    End Select

err:
End Sub



