VERSION 5.00
Begin VB.UserControl Graph 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000D&
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7830
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4875
   ScaleWidth      =   7830
   Begin VB.PictureBox picToPrinter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   4200
      ScaleHeight     =   555
      ScaleWidth      =   1005
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picToPrinterLegend 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1320
      ScaleHeight     =   555
      ScaleWidth      =   1005
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
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
      Height          =   1755
      Left            =   960
      ScaleHeight     =   1755
      ScaleWidth      =   3510
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3510
   End
End
Attribute VB_Name = "Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020

Private WithEvents mobjPoints      As Points
Attribute mobjPoints.VB_VarHelpID = -1

Private mudtControlProps    As gtypControlProps
Private mudtGraphProps      As gtypGraphProps

Private mblnDesignMode  As Boolean

Public Enum eBorderStyle
   egrNone = 0
   egrFixedSingle = 1
End Enum

Public Enum eAppearance
   egrFlat = 0
   egr3D = 1
End Enum

Private Type mtypPOINT
    X   As Long
    Y   As Long
    Index As Integer
End Type

Private Type mtypRECT
    Left    As Long
    Right   As Long
    Top     As Long
    Bottom  As Long
End Type



Public Enum LegendPrintConstants            'the enumerated for legend printing
    legPrintNone = 0
    legPrintGraph
    legPrintText
End Enum

Private uLegendPrintMode As LegendPrintConstants

Public Enum PrinterFitConstants             'the enumerated for printing
    prtFitCentered = 0
    prtFitStretched
    prtFitTopLeft
    prtFitTopRight
    prtFitBottomLeft
    prtFitBottomRight
End Enum

Private uPrinterFit As PrinterFitConstants
Private uPrinterOrientation As PrinterObjectConstants

Private bLegendAdded      As Boolean
Private bLegendClicked    As Boolean
Private bDisplayLegend    As Boolean
Private bResize           As Boolean
Private bResizeLegend     As Boolean


Private PointValue() As String
Private TestNumber() As String
Private PointIndex() As Integer
Private XMaxDevST As Integer
Private XMinST As Integer
Private XMaxST As Integer
Private MySTDMin As Double
Private MySTDMax As Double
Private MyMeanValue As Double

Private OffLeft As Long

Private Sub UserControl_Initialize()
    picDraw.FillStyle = vbFSSolid
    Set mobjPoints = New Points
End Sub

Private Sub UserControl_InitProperties()
    InitProperties
    mblnDesignMode = Not UserControl.Ambient.UserMode
    DrawControl
    PropertyChanged PB_STATE
End Sub

Private Sub UserControl_Paint()
    DrawGraph
End Sub

Private Sub UserControl_Terminate()
    Set mobjPoints = Nothing
    DoEvents
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    State = PropBag.ReadProperty(PB_STATE, State)
    mblnDesignMode = Not UserControl.Ambient.UserMode
    DrawControl
    Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PB_STATE, State
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    With UserControl
        picDraw.Move 0, 0, .ScaleWidth, .ScaleHeight
    End With
    Refresh
End Sub

Private Sub mobjPoints_Changed()
Static blnWorking As Boolean
    If Not blnWorking Then
        blnWorking = True
        RemovePoints
        If Not mblnDesignMode Then
            DrawGraph
        End If
        blnWorking = False
    End If
End Sub

Private Property Let GraphState(ByRef Value() As Byte)
Dim udtData     As gtypGraphData
    udtData.Data = Value
    LSet mudtGraphProps = udtData
End Property

Private Property Get GraphState() As Byte()
Dim udtData     As gtypGraphData
    LSet udtData = mudtGraphProps
    GraphState = udtData.Data
End Property


Friend Property Let ControlState(ByRef Value() As Byte)
Dim udtData     As gtypControlData
    udtData.Data = Value
    LSet mudtControlProps = udtData
End Property

Friend Property Get ControlState() As Byte()
Dim udtData     As gtypControlData
    LSet udtData = mudtControlProps
    ControlState = udtData.Data
End Property

Private Property Let State(ByRef Value() As Byte)
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .Contents = Value
        ControlState = .ReadProperty(PB_CONTROL)
        GraphState = .ReadProperty(PB_GRAPH)
    End With
    Set objPB = Nothing
End Property

Private Property Get State() As Byte()
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty PB_CONTROL, ControlState
        .WriteProperty PB_GRAPH, GraphState
        State = .Contents
    End With
    Set objPB = Nothing
End Property

Friend Property Let SuperState(ByRef Value() As Byte)
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .Contents = Value
        State = .ReadProperty(PB_STATE, State)
        mobjPoints.SuperState = .ReadProperty(PB_POINTS, mobjPoints.SuperState)
    End With
    Set objPB = Nothing
End Property

Friend Property Get SuperState() As Byte()
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty PB_STATE, State
        .WriteProperty PB_POINTS, mobjPoints.SuperState
        SuperState = .Contents
    End With
    Set objPB = Nothing
End Property

Friend Property Let FileState(ByRef Value() As Byte)
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .Contents = Value
        GraphState = .ReadProperty(PB_GRAPH, GraphState)
        mobjPoints.SuperState = .ReadProperty(PB_POINTS, mobjPoints.SuperState)
    End With
    Set objPB = Nothing
End Property

Friend Property Get FileState() As Byte()
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty PB_GRAPH, GraphState
        .WriteProperty PB_POINTS, mobjPoints.SuperState
        FileState = .Contents
    End With
    Set objPB = Nothing
End Property


Private Sub InitProperties()
    With mudtGraphProps
        .BackColor = RGB(255, 255, 255)
        .LineColor = RGB(84, 131, 169)
        .BarColor = vbColorBlueProgram  ' &HDEA68D
        .PointColor = &HFFFFC0
        .AxisColor = RGB(0, 0, 0)
        .GridColor = RGB(223, 223, 223)
        .FixedPoints = 20
        .XGridInc = 1
        .YGridInc = 10
        .MaxValue = 100
        .s1 = MaxValue / 3
        .s2 = MaxValue / 2
        .MinValue = 0
        .FadeIn = False
        .ShowGrid = True
        .ShowAxis = False
        .ShowLines = True
        .ShowPoints = True
        .ShowBars = True
        .BarWidth = 0.8
    End With
    With mudtControlProps
        .Redraw = True
        .BorderStyle = eBorderStyle.egrFixedSingle
        .Appearance = eAppearance.egr3D
    End With
    
    uPrinterFit = prtFitStretched ' prtFitCentered
    uPrinterOrientation = vbPRORPortrait ' vbPRDPVertical  ' vbPRORLandscape
    
    
End Sub

Public Property Get Points() As Points
    Set Points = mobjPoints
End Property

Public Property Let Redraw(ByVal Value As Boolean)
    mudtControlProps.Redraw = Value
    If Value Then
        Refresh
    End If
End Property

Public Property Get Redraw() As Boolean
    Redraw = mudtControlProps.Redraw
End Property



Public Property Let Appearance(ByVal Value As eAppearance)
    mudtControlProps.Appearance = Value
    UserControl.Appearance = Value
    DrawControl
    PropertyChanged PB_STATE
End Property

Public Property Get Appearance() As eAppearance
    Appearance = mudtControlProps.Appearance
End Property

Public Property Let BorderStyle(ByVal Value As eBorderStyle)
    mudtControlProps.BorderStyle = Value
    UserControl.BorderStyle = Value
    DrawControl
    PropertyChanged PB_STATE
End Property

Public Property Get BorderStyle() As eBorderStyle
    BorderStyle = mudtControlProps.BorderStyle
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.BackColor = Value
   ' DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mudtGraphProps.BackColor
End Property

Public Property Let LineColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.LineColor = Value
   ' DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get LineColor() As OLE_COLOR
    LineColor = mudtGraphProps.LineColor
End Property

Public Property Let BarColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.BarColor = Value
   ' DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get BarColor() As OLE_COLOR
    BarColor = mudtGraphProps.BarColor
End Property

Public Property Let PointColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.PointColor = Value
   ' DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get PointColor() As OLE_COLOR
    PointColor = mudtGraphProps.PointColor
End Property

Public Property Let AxisColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.AxisColor = Value
   ' DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get AxisColor() As OLE_COLOR
    AxisColor = mudtGraphProps.AxisColor
End Property
Public Property Let NumberColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.NumberColor = Value
   ' DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get NumberColor() As OLE_COLOR
    NumberColor = mudtGraphProps.NumberColor
End Property
Public Property Let MeanColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.MeanColor = Value
   ' DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get MeanColor() As OLE_COLOR
    MeanColor = mudtGraphProps.MeanColor
End Property
Public Property Let GridColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.GridColor = Value
   ' DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get GridColor() As OLE_COLOR
    GridColor = mudtGraphProps.GridColor
End Property
Public Property Let s1(ByVal Value As Double)
    mudtGraphProps.s1 = Value
End Property

Public Property Get s1() As Double
    s1 = mudtGraphProps.s1
End Property
Public Property Let s2(ByVal Value As Double)
    mudtGraphProps.s2 = Value
End Property

Public Property Get s2() As Double
    s2 = mudtGraphProps.s2
End Property

Public Property Let FixedPoints(ByVal Value As Long)
    mudtGraphProps.FixedPoints = Value
    RemovePoints
   ' DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get FixedPoints() As Long
    FixedPoints = mudtGraphProps.FixedPoints
End Property

Public Property Let XGridInc(ByVal Value As Long)
    mudtGraphProps.XGridInc = Value
   ' DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get XGridInc() As Long
    XGridInc = mudtGraphProps.XGridInc
End Property

Public Property Let YGridInc(ByVal Value As Double)
    mudtGraphProps.YGridInc = Value
   ' DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get YGridInc() As Double
    YGridInc = mudtGraphProps.YGridInc
End Property
Public Property Let MeanValue(ByVal Value As Double)
    mudtGraphProps.MeanValue = Value
    PropertyChanged PB_STATE
End Property

Public Property Get MeanValue() As Double
    MeanValue = mudtGraphProps.MeanValue
End Property
Public Property Let STDValue(ByVal Value As Double)
    mudtGraphProps.STDValue = Value
    PropertyChanged PB_STATE
End Property

Public Property Get STDValue() As Double
    STDValue = mudtGraphProps.STDValue
End Property



Public Property Let MaxValue(ByVal Value As Double)
    mudtGraphProps.MaxValue = Value
    PropertyChanged PB_STATE
End Property

Public Property Get MaxValue() As Double
    MaxValue = mudtGraphProps.MaxValue
End Property
Public Property Let X0Value(ByVal Value As Double)
    mudtGraphProps.X0Value = Value
    PropertyChanged PB_STATE
End Property

Public Property Get X0Value() As Double
    X0Value = mudtGraphProps.X0Value
End Property

Public Property Let IntVirgola(ByVal Value As Double)
    mudtGraphProps.IntVirgola = Value
    PropertyChanged PB_STATE
End Property

Public Property Get IntVirgola() As Double
    IntVirgola = mudtGraphProps.IntVirgola
End Property

Public Property Let devst(ByVal Value As Double)
    mudtGraphProps.devst = Value
    PropertyChanged PB_STATE
End Property

Public Property Get devst() As Double
    devst = mudtGraphProps.devst
End Property

Public Property Let YMaxDevST(ByVal Value As Double)
    mudtGraphProps.YMaxDevST = Value
    PropertyChanged PB_STATE
End Property

Public Property Get YMaxDevST() As Double
    YMaxDevST = mudtGraphProps.YMaxDevST
End Property

          
Public Property Let MinValue(ByVal Value As Double)
    mudtGraphProps.MinValue = Value
    PropertyChanged PB_STATE
End Property

Public Property Get MinValue() As Double
    MinValue = mudtGraphProps.MinValue
End Property

Public Property Let ShowGrid(ByVal Value As Boolean)
    mudtGraphProps.ShowGrid = Value
   ' DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get ShowGrid() As Boolean
    ShowGrid = mudtGraphProps.ShowGrid
End Property

Public Property Let ShowAxis(ByVal Value As Boolean)
    mudtGraphProps.ShowAxis = Value
    PropertyChanged PB_STATE
End Property

Public Property Get ShowAxis() As Boolean
    ShowAxis = mudtGraphProps.ShowAxis
End Property

Public Property Let ShowLines(ByVal Value As Boolean)
    mudtGraphProps.ShowLines = Value
    PropertyChanged PB_STATE
End Property

Public Property Get ShowLines() As Boolean
    ShowLines = mudtGraphProps.ShowLines
End Property

Public Property Let ShowGaussian(ByVal Value As Boolean)
    mudtGraphProps.ShowGaussian = Value
   ' DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get ShowGaussian() As Boolean
    ShowGaussian = mudtGraphProps.ShowGaussian
End Property


Public Property Let ShowBarsColor(ByVal Value As Boolean)
    mudtGraphProps.ShowBarsColor = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get ShowBarsColor() As Boolean
    ShowBarsColor = mudtGraphProps.ShowBarsColor
End Property

Public Property Let ShowBars(ByVal Value As Boolean)
    mudtGraphProps.ShowBars = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get ShowBars() As Boolean
    ShowBars = mudtGraphProps.ShowBars
End Property

Public Property Let ShowPoints(ByVal Value As Boolean)
    mudtGraphProps.ShowPoints = Value
    'DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get ShowPoints() As Boolean
    ShowPoints = mudtGraphProps.ShowPoints
End Property

Public Property Let ShowValue(ByVal Value As Boolean)
    mudtGraphProps.ShowValue = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get ShowValue() As Boolean
    ShowValue = mudtGraphProps.ShowValue
End Property
Public Property Let PrintPicture(ByVal Value As Boolean)
    If Value Then
       ' mnuMainPrint_Click
    End If
End Property

Public Property Get PrintPicture() As Boolean
    PrintPicture = True
End Property
Public Property Let FadeIn(ByVal Value As Boolean)
    mudtGraphProps.FadeIn = Value
   ' DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get FadeIn() As Boolean
    FadeIn = mudtGraphProps.FadeIn
End Property

Public Property Let BarWidth(ByVal Value As Single)
    mudtGraphProps.BarWidth = Value
   ' DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get BarWidth() As Single
    BarWidth = mudtGraphProps.BarWidth
End Property

Private Sub AddDefaultPoints()
    mobjPoints.Clear

End Sub



Private Sub RemovePoints()
    Do While mudtGraphProps.FixedPoints > 0 And mobjPoints.Count > mudtGraphProps.FixedPoints
        mobjPoints.Remove 1
    Loop
End Sub

Public Sub Refresh()
    DrawGraph
End Sub

Public Sub DrawControl()
    With UserControl
        .Appearance = mudtControlProps.Appearance
        .BorderStyle = mudtControlProps.BorderStyle
    End With
End Sub

Private Sub DrawGraph()
    Select Case mudtGraphProps.ShowGaussian
        Case True
            DrawGaussGraph
        Case Else
            DrawMeanGraph
    End Select
    


End Sub




Private Function GetPoints(ByVal plngLeft As Long, ByVal plngRight As Long, ByVal plngTop As Long, ByVal plngBottom As Long, ByVal plngBarWidth As Long) As mtypPOINT()
Dim udtPoints() As mtypPOINT
Dim lngCount    As Long
Dim lngIndex    As Long
Dim objPoint    As Point
Dim lngX        As Long
Dim lngPtCount  As Long
Dim lngFixedCount   As Long
Dim lngYAxis    As Long
    lngCount = mobjPoints.Count
    plngBarWidth = plngBarWidth + OffLeft
    If mudtGraphProps.ShowGaussian Then
        plngBarWidth = OffLeft
    Else
       ' Debug.Print "ciao"
    End If
    If lngCount > 0 Then
        If mudtGraphProps.FixedPoints = 0 Or lngCount < mudtGraphProps.FixedPoints Then
            lngPtCount = lngCount
            If mudtGraphProps.FixedPoints > 0 Then
                lngFixedCount = mudtGraphProps.FixedPoints
            Else
                lngFixedCount = lngCount
                If mudtGraphProps.ShowGaussian Then lngFixedCount = 4001
                
            End If
        Else
            lngPtCount = mudtGraphProps.FixedPoints
            lngFixedCount = mudtGraphProps.FixedPoints
        End If
        
        ReDim udtPoints(lngPtCount) As mtypPOINT
        ReDim PointValue(lngPtCount) As String
        ReDim TestNumber(lngPtCount) As String
        ReDim PointIndex(lngPtCount) As Integer
        
        For Each objPoint In mobjPoints
            lngIndex = lngIndex + 1
            If mudtGraphProps.FixedPoints > 0 And lngIndex > mudtGraphProps.FixedPoints Then
                Set objPoint = Nothing
                Exit For
            End If

            If lngIndex = 1 Then
                If lngFixedCount = 1 Then
                    lngX = plngLeft + (((plngRight - plngLeft)) / 2)
                Else
                    lngX = plngLeft + (plngBarWidth / 2)
                End If
            ElseIf lngIndex = lngFixedCount Then
                lngX = plngRight - (plngBarWidth / 2)
            Else
                lngX = (lngIndex - 1) * (((plngRight - plngLeft) - plngBarWidth) / (lngFixedCount - 1)) + (plngBarWidth / 2)
            End If

            udtPoints(lngIndex).X = lngX
            PointValue(lngIndex) = objPoint.RealValue
            PointIndex(lngIndex) = objPoint.Index
            TestNumber(lngIndex) = "T : " & objPoint.TestNumber
            
            If (mudtGraphProps.MaxValue - mudtGraphProps.MinValue) <> 0 Then
                lngYAxis = ((picDraw.ScaleHeight - 15) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.MaxValue
                
                ' trovo la Y del Punto....
                udtPoints(lngIndex).Y = lngYAxis - objPoint.Value * ((plngBottom - plngTop) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue))
                
                
                
                If FormatNumber(objPoint.Value, 7) = FormatNumber(mudtGraphProps.YMaxDevST, 7) Then
                    XMaxDevST = udtPoints(lngIndex).X
                End If
                
                If mudtGraphProps.s1 = 0 Then
                
                    If (objPoint.RealValue) = (mudtGraphProps.s1) Then
                        XMinST = udtPoints(lngIndex).X
                    End If
                    
                Else
                
                    If FormatNumber(objPoint.RealValue, mudtGraphProps.IntVirgola + 1) = FormatNumber(mudtGraphProps.s1, mudtGraphProps.IntVirgola + 1) Then
                        XMinST = udtPoints(lngIndex).X
                    End If
                    
                End If
                
                
               ' Debug.Print FormatNumber(objPoint.RealValue * 10000000, 5), FormatNumber(mudtGraphProps.s2 * 10000000, 5)
                If FormatNumber(objPoint.RealValue, mudtGraphProps.IntVirgola + 1) = FormatNumber(mudtGraphProps.s2, mudtGraphProps.IntVirgola + 1) Then
                    XMaxST = udtPoints(lngIndex).X
                End If
                
                               
                
                
            End If
          
        Next objPoint
                
     
          
    End If
    GetPoints = udtPoints
    
End Function

Private Sub DrawLine(ByRef pudtPt1 As mtypPOINT, ByRef pudtPt2 As mtypPOINT, ByVal plngColor As String)
    picDraw.Line (pudtPt1.X, pudtPt1.Y)-(pudtPt2.X, pudtPt2.Y), plngColor
End Sub

Private Sub DrawPoint(ByRef pudtPt As mtypPOINT, ByVal plngColor As Long, ByVal pstrCaption As String)
    picDraw.FillColor = plngColor
    picDraw.Circle (pudtPt.X, pudtPt.Y), 40, 0 'plngColor
End Sub

Private Sub DrawBar(ByRef pudtRect As mtypRECT, ByVal plngColor As Long)
    picDraw.FillColor = plngColor
    With pudtRect
        picDraw.Line (.Left, .Top)-(.Right, .Bottom), vbTimBlue, B
        
    End With
End Sub

Private Sub DrawGrid(ByRef pudtRect As mtypRECT, ByVal plngColor As Long, ByVal plngBarWidth As Long)
Dim lngCount    As Long
Dim lngIndex    As Long
Dim lngX        As Long
Dim lngY        As Long
Dim lngFixedCount   As Long
Dim lngYAxis    As Long
Dim lngStepY    As Long
Dim lngHeight   As Long
    lngCount = mobjPoints.Count
    If lngCount > 0 And mudtGraphProps.ShowGrid Then

        lngHeight = picDraw.ScaleHeight - 15

        If mudtGraphProps.FixedPoints = 0 Or lngCount < mudtGraphProps.FixedPoints Then
            If mudtGraphProps.FixedPoints > 0 Then
                lngFixedCount = mudtGraphProps.FixedPoints
            Else
                lngFixedCount = lngCount
            End If
        Else
            lngFixedCount = mudtGraphProps.FixedPoints
        End If
        For lngIndex = 1 To lngFixedCount
            If lngIndex = 1 Then
                If lngFixedCount = 1 Then
                    lngX = pudtRect.Left + (((pudtRect.Right - pudtRect.Left)) / 2)
                Else
                    lngX = pudtRect.Left + (plngBarWidth / 2)
                End If
            ElseIf lngIndex = lngFixedCount Then
                lngX = pudtRect.Right - (plngBarWidth / 2)
            Else
                lngX = (lngIndex - 1) * (((pudtRect.Right - pudtRect.Left) - plngBarWidth) / (lngFixedCount - 1)) + (plngBarWidth / 2)
            End If
            If lngIndex Mod mudtGraphProps.XGridInc = 0 Then
                picDraw.Line (lngX, 0)-(lngX, lngHeight), plngColor
            End If
        Next lngIndex

        'draw horizontal lines
        If (mudtGraphProps.MaxValue - mudtGraphProps.MinValue) > 0 Then
            lngYAxis = ((picDraw.ScaleHeight - 15) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.MaxValue
            lngStepY = (lngHeight / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.YGridInc
            For lngY = lngYAxis To 0 Step -lngStepY
                picDraw.Line (0, lngHeight - lngY)-(picDraw.ScaleWidth - 15, lngHeight - lngY), mudtGraphProps.GridColor
            Next lngY

            For lngY = lngYAxis To lngHeight Step lngStepY
                picDraw.Line (0, lngHeight - lngY)-(picDraw.ScaleWidth - 15, lngHeight - lngY), mudtGraphProps.GridColor
            Next lngY
        End If
    End If
End Sub

Public Sub SaveSettings(ByVal fileName As String)
    If Len(fileName) > 0 Then
        If Dir(fileName) <> vbNullString Then
            Kill fileName
        End If
    End If
    SaveFile fileName, FileState
End Sub

Public Sub LoadSettings(ByVal fileName As String)
    FileState = GetFile(fileName)
    Refresh
End Sub


Private Sub DrawMeanGraph()





Dim lngX        As Long
Dim lngY        As Long
Dim lngCount    As Long
Dim lngStepX    As Long
Dim lngStepY    As Long
Dim lngWidth    As Long
Dim lngHeight   As Long
Dim lngIndex    As Long
Dim udtPoints() As mtypPOINT
Dim lngYAxis    As Long
Dim lngBarWidth As Long
Dim lngFixedCount   As Long
Dim udtBar      As mtypRECT
Dim udtGrid     As mtypRECT

Dim STDsU As Long
Dim STDs3U As Long
Dim STDsD As Long
Dim STDs3D As Long

    On Error Resume Next
    If UserControl.Height > 0 And UserControl.Width > 0 Then
    If mudtControlProps.Redraw Or mblnDesignMode Then
    
    OffLeft = 600
        
    
        If mblnDesignMode Then
            AddDefaultPoints
        End If
        With picDraw
            .Cls
            .BackColor = mudtGraphProps.BackColor

            lngWidth = .ScaleWidth - 15 - OffLeft
            lngHeight = .ScaleHeight - 15
            
            's1 = mudtGraphProps.s1
            'S1Più = -mudtGraphProps.s1
            
            's2 = mudtGraphProps.s2
            'S2Più = -mudtGraphProps.s2
            
            'draw grid
            lngCount = mobjPoints.Count
            If mudtGraphProps.FixedPoints = 0 Or lngCount < mudtGraphProps.FixedPoints Then
                If mudtGraphProps.FixedPoints > 0 Then
                    lngFixedCount = mudtGraphProps.FixedPoints
                Else
                    lngFixedCount = lngCount
                End If
            Else
                lngFixedCount = mudtGraphProps.FixedPoints
            End If
            If lngFixedCount > 0 Then
                If mudtGraphProps.ShowBars Then
                    If lngCount > lngFixedCount Or Not mudtGraphProps.FadeIn Then
                        lngBarWidth = CLng((lngWidth / lngFixedCount) * mudtGraphProps.BarWidth)
                    ElseIf lngCount > 0 Then
                        lngBarWidth = CLng((lngWidth / lngCount) * mudtGraphProps.BarWidth)
                    End If
                End If
            End If

            udtPoints = GetPoints(0, lngWidth, 0, lngHeight, lngBarWidth)

            With udtGrid
                .Left = OffLeft
                .Top = 0
                .Right = lngWidth
                .Bottom = lngHeight
            End With
           ' DrawGrid udtGrid, mudtGraphProps.GridColor, lngBarWidth


            If (mudtGraphProps.MaxValue - mudtGraphProps.MinValue) > 0 Then
                lngYAxis = ((picDraw.ScaleHeight - 15) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.MaxValue
                
              
                MySTDMin = lngYAxis - mudtGraphProps.s1 * ((lngHeight) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue))
                MySTDMax = lngYAxis - mudtGraphProps.s2 * ((lngHeight) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue))
                MyMeanValue = lngYAxis - mudtGraphProps.MeanValue * ((lngHeight) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue))
        
        
                
                STDsU = lngHeight / 2 - (lngHeight / 2 - MySTDMax) / 2
                STDsD = lngHeight / 2 + (lngHeight / 2 - MySTDMax) / 2
                
                STDs3U = lngHeight / 2 - (lngHeight / 2 - MySTDMax) * 1.5
                STDs3D = lngHeight / 2 + (lngHeight / 2 - MySTDMax) * 1.5
                
            End If
            
picDraw.DrawStyle = 0



            picDraw.Font = "Calibri"

            picDraw.FontBold = True
            picDraw.ForeColor = mudtGraphProps.AxisColor 'mudtGraphProps.MeanColor
            
            
            
            picDraw.CurrentX = 200
            picDraw.CurrentY = lngHeight / 2 - 100
            picDraw.Print "STD"
            
            If MySTDMin <> lngHeight / 2 Then
        
                picDraw.ForeColor = &H4080&
                picDraw.CurrentX = 200
                picDraw.CurrentY = MySTDMin - 100
                picDraw.Print "Min"
                 picDraw.Line (OffLeft, MySTDMin)-(lngWidth + OffLeft / 2, MySTDMin), &H4080& ' mudtGraphProps.MeanColor     ' vbGreen ' mudtGraphProps.AxisColor
            End If
            If MySTDMax <> lngHeight / 2 And MySTDMax > 0 Then
                picDraw.ForeColor = &H4080&
                picDraw.CurrentX = 200
                picDraw.CurrentY = MySTDMax - 100
                picDraw.Print "Max"
                picDraw.Line (OffLeft, MySTDMax)-(lngWidth + OffLeft / 2, MySTDMax), &H4080& ' vbcolorred ' mudtGraphProps.AxisColor
            End If

          
  picDraw.DrawStyle = 2
             If Int(STDsU / 10) <> Int(lngHeight / 2 / 10) And Int(STDsU / 10) <> Int(MyMeanValue / 10) Then
                picDraw.ForeColor = mudtGraphProps.AxisColor  ' &H4080&
                picDraw.CurrentX = 270
                picDraw.CurrentY = STDsU - 100
                picDraw.Print "+s"
                picDraw.Line (OffLeft, STDsU)-(lngWidth + OffLeft / 2, STDsU), mudtGraphProps.AxisColor
            End If
          
              If Int(STDsD / 10) <> Int(lngHeight / 2 / 10) And Int(STDsD / 10) <> Int(MyMeanValue / 10) And MySTDMin <> lngHeight / 2 Then
                picDraw.ForeColor = mudtGraphProps.AxisColor  ' &H4080&
                picDraw.CurrentX = 270
                picDraw.CurrentY = STDsD - 100
                picDraw.Print "-s"
                picDraw.Line (OffLeft, STDsD)-(lngWidth + OffLeft / 2, STDsD), mudtGraphProps.AxisColor
            End If
            
            If Int(STDs3U / 10) <> Int(lngHeight / 2 / 10) And Int(STDs3U / 10) <> Int(MyMeanValue / 10) Then
                picDraw.ForeColor = mudtGraphProps.AxisColor  ' &H4080&
                picDraw.CurrentX = 220
                picDraw.CurrentY = STDs3U - 100
                picDraw.Print "+3s"
                picDraw.Line (OffLeft, STDs3U)-(lngWidth + OffLeft / 2, STDs3U), mudtGraphProps.AxisColor
            End If
          
              If Int(STDs3D / 10) <> Int(lngHeight / 2 / 10) And Int(STDs3D / 10) <> Int(MyMeanValue / 10) And MySTDMin <> lngHeight / 2 Then
                picDraw.ForeColor = mudtGraphProps.AxisColor  ' &H4080&
                picDraw.CurrentX = 220
                picDraw.CurrentY = STDs3D - 100
                picDraw.Print "-3s"
                picDraw.Line (OffLeft, STDs3D)-(lngWidth + OffLeft / 2, STDs3D), mudtGraphProps.AxisColor
            End If
 picDraw.DrawStyle = 1
             If MyMeanValue <> lngHeight / 2 Then
                picDraw.ForeColor = mudtGraphProps.MeanColor ' mudtGraphProps.AxisColor ' &H4080&
                picDraw.CurrentX = 100
                picDraw.CurrentY = MyMeanValue - 100
                picDraw.Print "Mean"
                picDraw.Line (OffLeft, MyMeanValue)-(lngWidth + OffLeft / 2, MyMeanValue), mudtGraphProps.MeanColor
            End If
                      
picDraw.DrawStyle = 0

'picDraw.DrawWidth = 1
           
             ' picDraw.Line (OffLeft, S1Più)-( lngWidth, S1Più), mudtGraphProps.MeanColor    ' vbGreen ' mudtGraphProps.AxisColor
            
             ' picDraw.Line (OffLeft, S2Più)-( lngWidth, S2Più), &H4080&   ' vbcolorred ' mudtGraphProps.AxisColor
             
             
             
             
                         'draw axis
            If mudtGraphProps.ShowAxis Then
            
                picDraw.DrawWidth = IIf(mudtGraphProps.ShowBars, 1, 2)
                picDraw.Line (OffLeft, 100)-(OffLeft, lngHeight - 300), mudtGraphProps.AxisColor
                picDraw.Line (OffLeft, lngHeight / 2)-(lngWidth + OffLeft / 2, lngHeight / 2), vbColorBlueProgram  ' mudtGraphProps.AxisColor
                picDraw.DrawWidth = 1
                picDraw.Line (OffLeft, 100)-(lngWidth + OffLeft / 2, 100), mudtGraphProps.AxisColor
                
                picDraw.Line (lngWidth + OffLeft / 2, 100)-(lngWidth + OffLeft / 2, lngHeight - 300), mudtGraphProps.AxisColor
                
                picDraw.Line (OffLeft, lngHeight - 300)-(lngWidth + OffLeft / 2, lngHeight - 300), mudtGraphProps.AxisColor
            End If
             
             
             
             
             
             
             
             
             
             
             
             
             
             
             
            'drawlines and bars
            If lngCount > 0 And (mudtGraphProps.ShowLines Or mudtGraphProps.ShowBars) Then
                For lngIndex = 1 To UBound(udtPoints)
                     
                     
                     udtPoints(lngIndex).X = udtPoints(lngIndex).X + OffLeft / 2
                     
                    
                    picDraw.ForeColor = mudtGraphProps.MeanColor   ' mudtGraphProps.AxisColor ' &H4080&
                    picDraw.CurrentX = udtPoints(lngIndex).X - 100
                    picDraw.CurrentY = lngHeight - 250
                    picDraw.Print lngIndex
                    
picDraw.DrawStyle = 2
                    picDraw.Line (udtPoints(lngIndex).X, 200)-(udtPoints(lngIndex).X, lngHeight - 300), mudtGraphProps.MeanColor 'mudtGraphProps.AxisColor
picDraw.DrawStyle = 0
                     
                    If mudtGraphProps.ShowBars Then
                     
                        If mudtGraphProps.ShowBarsColor Then
                            Select Case PointIndex(lngIndex)
                                Case 0
                                     mudtGraphProps.BarColor = vbColorBlueProgram
                                
                                Case 1
                                     mudtGraphProps.BarColor = &HA65911
                                Case 2
                                     mudtGraphProps.BarColor = &HEBC99B
                                Case 3
                                    mudtGraphProps.BarColor = &H8000000D
                                Case Else
                                    mudtGraphProps.BarColor = vbColorBlueProgram
                            End Select
                        Else
                            mudtGraphProps.BarColor = vbColorBlueProgram
                        End If
                        
                        
                        
                        udtBar.Left = udtPoints(lngIndex).X - (lngBarWidth / 2)
                        udtBar.Right = udtPoints(lngIndex).X + (lngBarWidth / 2)

                        udtBar.Top = udtPoints(lngIndex).Y
                        udtBar.Bottom = lngYAxis
                        DrawBar udtBar, mudtGraphProps.BarColor
                    End If
                    If mudtGraphProps.ShowLines And lngIndex > 1 Then
                      picDraw.DrawWidth = IIf(mudtGraphProps.ShowBars, 1, 2)
                      DrawLine udtPoints(lngIndex - 1), udtPoints(lngIndex), mudtGraphProps.LineColor
                      picDraw.DrawWidth = 1
                    End If
                Next lngIndex
            End If

            

            'draw points
            If lngCount > 0 And mudtGraphProps.ShowPoints Then
                For lngIndex = 1 To UBound(udtPoints)
                    If lngIndex Mod mudtGraphProps.XGridInc = 0 Or lngIndex = 1 Then
                    
                        
                       
                        If udtPoints(lngIndex).Y < lngHeight / 2 Then
                            If udtPoints(lngIndex).Y < MySTDMax Then
                                 mudtGraphProps.PointColor = vbColorRed
                            Else
                                mudtGraphProps.PointColor = vbGreen '&H4000&
                            End If
                        Else
                            If udtPoints(lngIndex).Y > MySTDMin Then
                                 mudtGraphProps.PointColor = vbColorRed
                            Else
                                mudtGraphProps.PointColor = vbGreen
                            End If
                        
                        End If
                        
                        
                        picDraw.FontBold = True
                        picDraw.FontSize = 8
                        ' stampo i valori
                        
                        If mudtGraphProps.ShowValue Then
                            picDraw.ForeColor = mudtGraphProps.NumberColor ' &HA0A0A0
                            picDraw.CurrentX = udtPoints(lngIndex).X - 40 * Len(PointValue(lngIndex))
                            If udtPoints(lngIndex).Y < lngHeight / 2 Then
                            
                                ' SOPRA LA MEDIA
                                
                                picDraw.CurrentY = udtPoints(lngIndex).Y - 300
                                If picDraw.CurrentY < 50 Then
                                
                                    picDraw.CurrentY = udtPoints(lngIndex).Y + 300
                                    picDraw.ForeColor = mudtGraphProps.NumberColor ' mudtGraphProps.LineColor   ' &HA0A0A0
                                End If
                                
                              

                            Else
                                
                                ' SOTTO LA MEDIA
                                
                                 picDraw.CurrentY = udtPoints(lngIndex).Y + 100
                                If picDraw.CurrentY > lngHeight - 200 Then
                                      picDraw.CurrentY = udtPoints(lngIndex).Y - 400
                                    picDraw.ForeColor = mudtGraphProps.NumberColor ' mudtGraphProps.LineColor   ' &HA0A0A0
                                End If
                               
                            End If
                            picDraw.Print PointValue(lngIndex) 'udtPoints(lngIndex).Y
                            
                            
                            
                            ' stampa il Numero di Test
                            
                            picDraw.ForeColor = mudtGraphProps.NumberColor ' &HA0A0A0
                            picDraw.CurrentX = udtPoints(lngIndex).X - 25 * Len(TestNumber(lngIndex))
                            If udtPoints(lngIndex).Y < lngHeight / 2 Then
                            
                                ' SOPRA LA MEDIA
                                
                                picDraw.CurrentY = udtPoints(lngIndex).Y - 500
                                If picDraw.CurrentY < 50 Then
                                
                                    picDraw.CurrentY = udtPoints(lngIndex).Y + 500
                                    picDraw.ForeColor = mudtGraphProps.NumberColor ' mudtGraphProps.LineColor   ' &HA0A0A0
                                End If
                                
                              

                            Else
                                
                                ' SOTTO LA MEDIA
                                
                                 picDraw.CurrentY = udtPoints(lngIndex).Y + 300
                                If picDraw.CurrentY > lngHeight - 200 Then
                                      picDraw.CurrentY = udtPoints(lngIndex).Y - 600
                                    picDraw.ForeColor = mudtGraphProps.NumberColor ' mudtGraphProps.LineColor   ' &HA0A0A0
                                End If
                               
                            End If
                            picDraw.Print TestNumber(lngIndex) 'udtPoints(lngIndex).Y
                            
                        End If
                        
                        
                        
                        
                        
                         DrawPoint udtPoints(lngIndex), mudtGraphProps.PointColor, "Woof"
                        
                        
                    End If
                Next lngIndex
            End If
            


            'copy picture to usercontrol
            BitBlt UserControl.hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, SRCCOPY
        End With
    End If
    End If
End Sub








Private Sub DrawGaussGraph()





Dim lngX        As Long
Dim lngY        As Long
Dim lngCount    As Long
Dim lngStepX    As Long
Dim lngStepY    As Long
Dim lngWidth    As Long
Dim lngHeight   As Long
Dim lngIndex    As Long
Dim udtPoints() As mtypPOINT
Dim lngYAxis    As Long
Dim lngBarWidth As Long
Dim lngFixedCount   As Long
Dim udtBar      As mtypRECT
Dim udtGrid     As mtypRECT

Dim s1 As Long
Dim s2 As Long
Dim S1Più As Long
Dim S2Più As Long
Dim STDXValue As Double
Dim Virgola As Integer

    On Error Resume Next
    If UserControl.Height > 0 And UserControl.Width > 0 Then
    If mudtControlProps.Redraw Or mblnDesignMode Then
    
    OffLeft = 600
        ' MsgBox mudtGraphProps.MeanValue
    
    Virgola = mudtGraphProps.IntVirgola
    
        With picDraw
            .Cls
            .BackColor = mudtGraphProps.BackColor

            lngWidth = .ScaleWidth - 15 - OffLeft
            lngHeight = .ScaleHeight - 15

            'draw grid
            lngCount = mobjPoints.Count
            If mudtGraphProps.FixedPoints = 0 Or lngCount < mudtGraphProps.FixedPoints Then
                If mudtGraphProps.FixedPoints > 0 Then
                    lngFixedCount = mudtGraphProps.FixedPoints
                Else
                    lngFixedCount = lngCount
                End If
            Else
                lngFixedCount = mudtGraphProps.FixedPoints
            End If
            If lngFixedCount > 0 Then
                If mudtGraphProps.ShowBars Then
                    If lngCount > lngFixedCount Or Not mudtGraphProps.FadeIn Then
                        lngBarWidth = CLng((lngWidth / lngFixedCount) * mudtGraphProps.BarWidth)
                    ElseIf lngCount > 0 Then
                        lngBarWidth = CLng((lngWidth / lngCount) * mudtGraphProps.BarWidth)
                    End If
                End If
            End If
        Debug.Print mobjPoints.Count
            udtPoints = GetPoints(0, lngWidth, 0, lngHeight, lngBarWidth)

            With udtGrid
                .Left = OffLeft
                .Top = 0
                .Right = lngWidth
                .Bottom = lngHeight
            End With
           ' DrawGrid udtGrid, mudtGraphProps.GridColor, lngBarWidth

            If (mudtGraphProps.MaxValue - mudtGraphProps.MinValue) > 0 Then
                lngYAxis = ((picDraw.ScaleWidth - 15) / (1000)) * 1000
                
              
               
               ' MyMeanValue = lngWidth / 2 + OffLeft / 2
                'MyMeanValue = mudtGraphProps.MeanValue * (lngWidth / 2) / mudtGraphProps.X0Value + OffLeft / 2
                MyMeanValue = XMaxDevST + OffLeft / 2
                
                MySTDMin = XMinST + OffLeft / 2 '(mudtGraphProps.s1 * (lngWidth / 2) / mudtGraphProps.X0Value) + OffLeft / 2
                MySTDMax = XMaxST + OffLeft / 2 '(mudtGraphProps.s2 * (lngWidth / 2) / mudtGraphProps.X0Value) + OffLeft / 2
                
                'MySTDMin = (mudtGraphProps.s1 * XMaxDevST / mudtGraphProps.MeanValue) + OffLeft / 2
                'MySTDMax = (mudtGraphProps.s2 * XMaxDevST / mudtGraphProps.MeanValue) + OffLeft / 2
                
                If MySTDMin < OffLeft Then MySTDMin = OffLeft
                
                STDXValue = ((MySTDMax - MySTDMin) / 2) + MySTDMin '- OffLeft
             End If

 picDraw.DrawStyle = 0
            
            'draw axis
            If mudtGraphProps.ShowAxis Then
                picDraw.Line (OffLeft, 100)-(OffLeft, lngHeight - 300), mudtGraphProps.AxisColor
                If mudtGraphProps.MaxValue <= 0 Then
                    picDraw.Line (OffLeft, 0)-(lngWidth, 0), mudtGraphProps.AxisColor
                ElseIf mudtGraphProps.MinValue < 0 Then
                    picDraw.Line (OffLeft, (lngHeight / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.MaxValue)-(lngWidth + OffLeft, (lngHeight / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.MaxValue), mudtGraphProps.AxisColor
                Else
                    picDraw.Line (OffLeft, lngHeight - 300)-(lngWidth + OffLeft, lngHeight - 300), mudtGraphProps.AxisColor
                End If
            End If


           
 picDraw.DrawStyle = 2


            picDraw.Font = "Calibri"

            picDraw.FontBold = True
            picDraw.ForeColor = mudtGraphProps.AxisColor 'mudtGraphProps.MeanColor
            
            
            
            'picDraw.CurrentX = 200
            'picDraw.CurrentY = lngHeight / 2 - 100
            'picDraw.Print "STD"
        
            picDraw.ForeColor = &H4080&
            picDraw.CurrentX = MySTDMin - 100
            picDraw.CurrentY = lngHeight - 250
            picDraw.Print "Min"
            picDraw.CurrentY = 450
            picDraw.CurrentX = MySTDMin - 150
            picDraw.Print FormatNumber(mudtGraphProps.s1, Virgola)
            
            picDraw.Line (MySTDMin, 700)-(MySTDMin, lngHeight - 300), &H4080& ' mudtGraphProps.MeanColor     ' vbGreen ' mudtGraphProps.AxisColor
            
            If mudtGraphProps.s2 > 0 Then
                picDraw.ForeColor = &H4080&
                picDraw.CurrentX = MySTDMax - 100
                picDraw.CurrentY = lngHeight - 250
                picDraw.Print "Max"
                picDraw.CurrentY = 450
                picDraw.CurrentX = MySTDMax - 150 * (Len(CStr(mudtGraphProps.s2)) / 3)
                picDraw.Print FormatNumber(mudtGraphProps.s2, Virgola)
                
                picDraw.Line (MySTDMax, 700)-(MySTDMax, lngHeight - 300), &H4080&  ' vbcolorred ' mudtGraphProps.AxisColor
            
            End If
            
            If mudtGraphProps.MeanValue > 0 Then
                picDraw.ForeColor = mudtGraphProps.MeanColor ' mudtGraphProps.AxisColor ' &H4080&
                picDraw.CurrentX = MyMeanValue - 200
                picDraw.CurrentY = lngHeight - 250
                picDraw.Print "Mean"
                picDraw.CurrentY = 70
                picDraw.CurrentX = MyMeanValue - 150 * (Len(CStr(mudtGraphProps.MeanValue)) / 4)
                picDraw.Print FormatNumber(mudtGraphProps.MeanValue, Virgola)
                
                picDraw.Line (MyMeanValue, 300)-(MyMeanValue, lngHeight - 300), mudtGraphProps.MeanColor 'mudtGraphProps.AxisColor
                
                picDraw.ForeColor = mudtGraphProps.MeanColor  ' mudtGraphProps.AxisColor ' &H4080&
                picDraw.CurrentY = 700
                picDraw.CurrentX = MyMeanValue + 150
                picDraw.Print "StDev = " & FormatNumber(mudtGraphProps.devst, Virgola)
            End If
            
picDraw.DrawStyle = 2


            If Abs(STDXValue - MyMeanValue) > 200 And mudtGraphProps.s1 > 0 Then
                
                picDraw.ForeColor = mudtGraphProps.BarColor ' mudtGraphProps.AxisColor ' &H4080&
                picDraw.CurrentX = STDXValue - 150
                picDraw.CurrentY = lngHeight - 250
                picDraw.Print "STD"
                picDraw.CurrentY = 250
                picDraw.CurrentX = STDXValue - 250 * (Len(CStr(mudtGraphProps.STDValue)) / 4)
                picDraw.Print FormatNumber(mudtGraphProps.STDValue, Virgola)
                
                picDraw.Line (STDXValue, 500)-(STDXValue, lngHeight - 300), mudtGraphProps.BarColor
            End If
 
picDraw.DrawStyle = 0
            'drawlines and bars
            If lngCount > 0 And (mudtGraphProps.ShowLines Or mudtGraphProps.ShowBars) Then
                For lngIndex = 1 To UBound(udtPoints)
                     udtPoints(lngIndex).X = udtPoints(lngIndex).X + OffLeft / 2
                     udtPoints(lngIndex).Y = udtPoints(lngIndex).Y - 300
                    If mudtGraphProps.ShowBars Then
                     
                        If mudtGraphProps.ShowBarsColor Then
                            Select Case PointIndex(lngIndex)
                                Case 0
                                     mudtGraphProps.BarColor = vbColorBlueProgram
                                
                                Case 1
                                     mudtGraphProps.BarColor = &HA65911
                                Case 2
                                mudtGraphProps.BarColor = &HEBC99B
                                Case 3
                            mudtGraphProps.BarColor = &H8000000D
                            
                            End Select
                        Else
                            mudtGraphProps.BarColor = vbColorBlueProgram
                        End If
                        
                        
                        
                        udtBar.Left = udtPoints(lngIndex).X - (lngBarWidth / 2)
                        udtBar.Right = udtPoints(lngIndex).X + (lngBarWidth / 2)

                        udtBar.Top = udtPoints(lngIndex).Y
                        udtBar.Bottom = lngYAxis
                        DrawBar udtBar, mudtGraphProps.BarColor
                    End If
                    If mudtGraphProps.ShowLines And lngIndex > 1 Then
                        picDraw.DrawWidth = 2
                        DrawLine udtPoints(lngIndex - 1), udtPoints(lngIndex), vbColorBlueProgram 'mudtGraphProps.LineColor
                        picDraw.DrawWidth = 1
                  End If
                Next lngIndex
            End If


            'copy picture to usercontrol
            BitBlt UserControl.hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, SRCCOPY
        End With
    End If
    End If
End Sub

