Attribute VB_Name = "ResizeMe"
Option Explicit

Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type

Public m_ControlGridFontSize As Double
Public m_ControlGridRowHeight As Double
Public m_ControlGridColWidth As Double



Public m_ControlGridFontSizeOld As Double
Public m_ControlGridColWidthOld  As Double
Public m_ControlGridRowHeightOld As Double


Public m_ControlPositions() As ControlPositionType
Public m_FormWid As Single
Public m_FormHgt As Single
