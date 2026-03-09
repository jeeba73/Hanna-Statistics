Attribute VB_Name = "modTypes"
Option Explicit

Public Type gtypPointProps
    Value       As Double
    Index       As Integer
    RealValue   As Double
    TestNumber  As Integer
End Type

Public Type gtypPointData
    Data        As String * 4
End Type

Public Type gtypGraphProps
    BackColor       As Long
    LineColor       As Long
    BarColor        As Long
    PointColor      As Long
    AxisColor       As Long
    MeanColor       As Long
    NumberColor     As Long
    GridColor       As Long
    FixedPoints     As Long
    XGridInc        As Long
    YGridInc        As Double
    MaxValue        As Double
    MinValue        As Double
    STDValue        As Double
    MeanValue       As Double
    devst           As Double
    YMaxDevST       As Double
    X0Value         As Double
    IntVirgola      As Double
    s1              As Double
    s2              As Double
    FadeIn          As Boolean
    BarWidth        As Single
    ShowGrid        As Boolean
    ShowAxis        As Boolean
    ShowLines       As Boolean
    ShowBars        As Boolean
    ShowBarsColor   As Boolean
    ShowGaussian    As Boolean
    ShowPoints      As Boolean
    ShowValue       As Boolean
   ' Image           As Image
End Type

Public Type gtypGraphData
    Data    As String * 36
End Type

Public Type gtypControlProps
    Redraw      As Boolean
    Appearance  As Integer
    BorderStyle As Integer
End Type

Public Type gtypControlData
    Data        As String * 3
End Type
