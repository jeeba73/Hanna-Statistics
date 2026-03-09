Attribute VB_Name = "Grafico"
Public X0, Y0, X1, Y1 As Long
Public X_Axis_Min, X_Axis_Max, Y_Axis_Min, Y_Axis_Max As Double
Public X_Tic_Number, Y_Tic_Number As Integer
Public Pr_X, Pr_Y As Double
Public Cur_X, Cur_Y As Double






Public Sub InitGraph(Picture As PictureBox, X_Min, X_Max, Y_Min, Y_Max, X_Tic, Y_Tic, ByVal typeGrf As Integer, Optional StartTime As String, Optional StopTime As String)

Dim MyForeCOlor As Long         'Set graph's color. Can use QBColor or RGB() function
MyForeCOlor = RGB(0, 0, 0)      '
Dim passo As Double



'Y_Min = IIf(Y_Min / 1000 > 1, Y_Min / 1000, Y_Min)
'Y_Max = IIf(Y_Max / 1000 > 1, Y_Max / 1000, Y_Max)


X0 = 1000                       'Set the public variables
Y0 = Picture.Height - 1000         'for the graph window. X0,Y0,X1,Y1 are
X1 = Picture.Width - 400            'actual form coordinates in twips
Y1 = 300                        '
X_Axis_Min = X_Min              'Axis limits that may change according
X_Axis_Max = X_Max              'to the actual X,Y data values
Y_Axis_Min = Y_Min              'given later in PlotData() function
Y_Axis_Max = Y_Max
X_Tic_Number = X_Tic
Y_Tic_Number = Y_Tic

Picture.AutoRedraw = True                   'Need this so that a box is drawn when Picture loads.
'Picture.BackColor = RGB(255, 255, 255)
'Picture.ForeColor = MyForeCOlor
Picture.DrawStyle = 0                       'DrawStyle is solid line
Picture.Cls

Picture.Line (X0, Y0)-(X1, Y1), , B       'Draw a box


'Draw X-Tics
x_inc = ((X1 - X0) / X_Tic_Number)
If X_Axis_Max = 0 Then
    passo = 0
Else

    passo = (X1 - X0) / X_Axis_Max
End If
tic_val_x = (X_Axis_Max - X_Axis_Min) / X_Tic_Number
tic_y = Y1
Picture.DrawStyle = 2 'DrawStyle is now dots



        For i = 1 To X_Tic_Number + 1
        
            tic_x = X0 + (i - 1) * x_inc
            Picture.Line (tic_x, Y0)-(tic_x, tic_y)
            Picture.CurrentX = tic_x - 120
            Picture.CurrentY = Y0 + 200
            xxx = (X_Axis_Min + (i - 1) * tic_val_x)
            virgola = 0 'IIf(Int(xxx) = xxx, 0, 2)
            If typeGrf = 0 Or typeGrf = 2 Then
                Picture.Print Int(xxx)
            Else
               Picture.Print FormatNumber(xxx, virgola)
            End If
          '  Picture.CurrentX = tic_x - 120
          '  Picture.CurrentY = Y0 + 300
          '  If i = 1 Then Picture.Print StartTime
          '  If i = X_Tic_Number + 1 Then Picture.Print StopTime
        Next i
    

Picture.DrawStyle = 0

Pr_X = -100
Pr_Y = -100

End Sub


