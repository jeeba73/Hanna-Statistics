Attribute VB_Name = "Autocomplete"
Public Function AutoSel(Cmb As ComboBox, KeyCode As Integer)
       '   Debug.Print KeyCode
          If KeyCode = vbEnter Then Exit Function
          If KeyCode = 8 Then Exit Function    'Backspace
          If KeyCode = 37 Then Exit Function  'left key
          If KeyCode = 38 Then Exit Function 'up arrow key
          If KeyCode = 39 Then Exit Function  'right key
          If KeyCode = 40 Then Exit Function  'down arrow key
          If KeyCode = 46 Then Exit Function  'delete key
          If KeyCode = 33 Then Exit Function  'page up key
          If KeyCode = 34 Then Exit Function  'page down key
          If KeyCode = 35 Then Exit Function  'end key
          If KeyCode = 36 Then Exit Function  'home key
         
         
          Dim Text As String
          Text = Cmb.Text
         
          Dim i As Long
          Dim Temp As String
         
         On Error Resume Next
         
          For i = 0 To Cmb.ListCount
              Temp = Left(Cmb.List(i), Len(Text))
              If LCase(Temp) = LCase(Text) Then
                  Cmb.Text = Cmb.List(i)
                 ' Cmb.ListIndex = i
                  Cmb.SelStart = Len(Text)
                  Cmb.SelLength = Len(Cmb.List(i))
                  'Cmb.SetFocus
              End If
              
          Next
          
          
          On Error GoTo 0
          
          
      End Function



Public Function SearchCode(ByVal Code As String, ByVal Grid As Grid, ByVal bTutti As Boolean)
Dim i As Integer

Code = Trim(Code)

If Code = "" Then bTutti = True


With Grid
    If .Rows < 2 Then Exit Function
    .AutoRedraw = False
    For i = 1 To .Rows - 1
        If bTutti Then
            .RowHeight(i) = 25
        Else
        
            If InStr(UCase(.Cell(i, 1).Text), UCase(Code)) Then
                .RowHeight(i) = 25
            Else
                .RowHeight(i) = 0
            
            End If
        End If
    Next
    
    .Refresh
    .AutoRedraw = True
End With
End Function

