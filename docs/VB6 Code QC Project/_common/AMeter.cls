VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Use this class to create a Graphical Meter Display, such as Voice Meter, Graphic Eq. etc."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------------
' Class: AMeter
'
' Purpose: Use this class to create a Display
'          like Voice Meter, Graphic Eq etc.
'
' Author: Dipak Auddy.
' E-Mail: auddy@gawab.com
' Date  : Dec 05, 2004.
'
' Note(s):
'       (1) You will need a PictureBox.
'       (2) All measurements in Pixels.
'       (3) Range of Values should be
'           in 0->100 Range.
'       (4) First Set/Let the Surface,
'           then anything else......
'--------------------------------------------------
Option Explicit
'local variable(s) to hold property value(s)
Private mvarSurface         As PictureBox

Private mvarBarForeColor    As Long
Private mvarBarBackColor    As Long
Private mvarSurfaceColor    As Long

Private mvarNumBars         As Integer
Private mvarBarWidth        As Integer
Private mvarBarHeight       As Integer
Private mvarBarGap          As Integer
Private mvarXLeft           As Integer

Public Property Get BarBackColor() As Long
Attribute BarBackColor.VB_Description = "Bar Background Color"

    BarBackColor = mvarBarBackColor

End Property

Public Property Let BarBackColor(ByVal vData As Long)

    mvarBarBackColor = vData

End Property

Public Property Get BarForeColor() As Long
Attribute BarForeColor.VB_Description = "Bar Foreground Color"

    BarForeColor = mvarBarForeColor

End Property

Public Property Let BarForeColor(ByVal vData As Long)

    mvarBarForeColor = vData

End Property

Public Property Get BarGap() As Integer
Attribute BarGap.VB_Description = "Distance between two bars. (Default 3)"

    BarGap = mvarBarGap

End Property

Public Property Let BarGap(ByVal vData As Integer)

    mvarBarGap = vData

End Property

Public Property Get BarHeight() As Integer
Attribute BarHeight.VB_Description = "Height of the each Bar. (Default 1)"

    BarHeight = mvarBarHeight

End Property

Public Property Let BarHeight(ByVal vData As Integer)

    mvarBarHeight = vData

End Property

Public Property Get BarWidth() As Integer
Attribute BarWidth.VB_Description = "Bar Width. (Default 20)."

    BarWidth = mvarBarWidth

End Property

Public Property Let BarWidth(ByVal vData As Integer)

    mvarBarWidth = vData

End Property

Private Sub Class_Initialize()

'Default Bar...
'20 pix wide, 1 pix height, with 3 pix gap
'of total 10 bars
    
    mvarBarForeColor = &H4000&
    mvarBarBackColor = &H4000&
    mvarSurfaceColor = vbBlack
    mvarXLeft = 0
    mvarBarWidth = 100
    mvarBarHeight = 1
    mvarBarGap = 3
    mvarNumBars = 30

End Sub

Private Sub Class_Terminate()

    Set mvarSurface = Nothing

End Sub


Public Sub Draw(ByVal NewVal As Long, m_foreColor, m_backColor)

'NewVal should be in Range 0-100

On Error Resume Next
If NewVal < 0 Then NewVal = 0
Dim i As Integer

    For i = 0 To mvarNumBars * IIf(m_ControlGridRowHeight > 1, m_ControlGridRowHeight, 1)
        mvarSurface.Line (mvarXLeft, mvarXLeft + i * mvarBarGap)-(mvarXLeft + mvarBarWidth, Xleft + i * mvarBarGap + mvarBarHeight), IIf(NewVal >= 100 - i * 100 \ (mvarNumBars * IIf(m_ControlGridRowHeight > 1, m_ControlGridRowHeight, 1)) And NewVal <> 0, m_foreColor, m_backColor), BF
     ' mvarSurface.Line (mvarXLeft, mvarXLeft + i * mvarBarGap)-(mvarXLeft + mvarBarWidth, Xleft + i * mvarBarGap + mvarBarHeight), IIf(NewVal >= 100 - i * 100 \ (mvarNumBars) And NewVal <> 0, m_foreColor, m_backColor), BF
    Next i

End Sub
Public Property Get NumBars() As Integer
Attribute NumBars.VB_Description = "Numbers of Bers. (Default 10)"

    NumBars = mvarNumBars

End Property

Public Property Let NumBars(ByVal vData As Integer)

    mvarNumBars = vData

End Property

Public Property Let Surface(ByRef vData As PictureBox)
Attribute Surface.VB_Description = "Usually a PictureBox object on which Bar(s) are drawn."

'Usualy this should be Set, but I changed it to Let
'so no need to use the keyword 'Set' on your Form's code...

    Set mvarSurface = vData
    mvarSurface.ScaleMode = vbPixels
    mvarSurface.DrawWidth = 1
    mvarSurface.BackColor = mvarSurfaceColor
    

End Property

Public Property Get SurfaceColor() As Long
Attribute SurfaceColor.VB_Description = "Background Color of Surface (aka PictureBox)."

    SurfaceColor = mvarSurfaceColor

End Property

Public Property Let SurfaceColor(ByVal vData As Long)

    mvarSurfaceColor = vData
    mvarSurface.BackColor = mvarSurfaceColor

End Property

Public Property Get Xleft() As Integer
Attribute Xleft.VB_Description = "Left Margin. (Default 7)."

    Xleft = mvarXLeft

End Property

Public Property Let Xleft(ByVal vData As Integer)

    mvarXLeft = vData

End Property

':)Code Fixer V2.2.0 (26/04/2005 7:35:47 PM) 29 + 147 = 176 Lines Thanks Ulli for inspiration and lots of code.
