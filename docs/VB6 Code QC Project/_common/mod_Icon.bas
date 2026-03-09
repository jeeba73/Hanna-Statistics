Attribute VB_Name = "mod_Icon"
Option Explicit





Public Function Icona_MouseMove(ByVal Text As String, ByVal Icona As Image, ByVal Label As Label, ByVal FormWidth As Integer, Optional ByVal bValue As Boolean = True)
Dim iLeft As Integer
Dim iTop As Integer

    If Text = "" Then Exit Function

    Label.Caption = Text
    
    iLeft = Icona.Left - (Label.Width / 2 - Icona.Width / 2)
    If iLeft < 200 Then iLeft = 200 * m_ControlGridColWidth
    If iLeft > FormWidth - Label.Width - 200 Then iLeft = (FormWidth - Label.Width - 200) '* m_ControlGridColWidth
    iTop = Icona.TOp - (IIf(bValue, 860 * m_ControlGridRowHeight, -700 * m_ControlGridRowHeight))
    Label.Left = iLeft
    Label.TOp = iTop
    Label.Visible = True
    

End Function


Public Function Picture_MouseMove(ByVal Text As String, ByVal Picture As PictureBox, ByVal Label As Label, ByVal FormWidth As Integer, Optional ByVal bValue As Boolean = True)
Dim iLeft As Integer
Dim iTop As Integer

    If Text = "" Then Exit Function

    Label.Caption = Text
    
    iLeft = Picture.Left - (Label.Width / 2 - Picture.Width / 2)
    If iLeft < 200 Then iLeft = 200
    If iLeft > FormWidth - Label.Width - 200 Then iLeft = FormWidth - Label.Width - 200
    iTop = Picture.TOp - (IIf(bValue, 700, -800))
    Label.Left = iLeft
    Label.TOp = iTop
    Label.Visible = True
    

End Function

