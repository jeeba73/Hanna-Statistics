VERSION 5.00
Begin VB.UserControl ctlDateScroll 
   BackColor       =   &H00644603&
   BackStyle       =   0  'Transparent
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   ScaleHeight     =   690
   ScaleWidth      =   6105
   Begin VB.TextBox txtMain 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   630
      Index           =   0
      Left            =   3360
      TabIndex        =   2
      Text            =   "January"
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtMain 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   1
      Left            =   600
      TabIndex        =   0
      Text            =   "01/01/2004"
      Top             =   0
      Width           =   2655
   End
   Begin VB.VScrollBar scrMain 
      Height          =   645
      Index           =   0
      Left            =   0
      Min             =   -32600
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   20
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.VScrollBar scrMain 
      Height          =   645
      Index           =   1
      Left            =   5520
      Min             =   -32600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   20
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "ctlDateScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_DateMode As eDateScrollMode
Dim m_Locked As Boolean

Public Event Change()
Public Event Click()

Private Sub scrMain_Change(Index As Integer)
    Dim nStart As Integer
    Dim nLength As Integer
    Dim nField As String
    Dim sMonth As String
    Dim sYear As String
    
    If scrMain(Index).Value = 0 Or (IsDate(txtMain(Index).Text) = False And (m_DateMode = dsDate Or m_DateMode = dsTime)) Then
        If scrMain(Index).Value <> 0 Then
            scrMain(Index).Value = 0
        End If
        Exit Sub
    ElseIf m_DateMode = dsYear Then
        If Len(Trim(txtMain(Index).Text)) <> 2 And Len(Trim(txtMain(Index).Text)) <> 4 Then
            If scrMain(Index).Value <> 0 Then
                scrMain(Index).Value = 0
            End If
            Exit Sub
        End If
    ElseIf m_DateMode = dsMonth Then
        If IsDate(Trim(txtMain(Index).Text) & " 01,2000") = False Then
            If scrMain(Index).Value <> 0 Then
                scrMain(Index).Value = 0
            End If
            Exit Sub
        End If
    End If
    
    Select Case m_DateMode
        Case dsDate
            nLength = 2
            If txtMain(Index).SelStart < 3 Then
                nStart = 0
                nField = "m"
            ElseIf txtMain(Index).SelStart >= 3 And txtMain(Index).SelStart < 6 Then
                nStart = 3
                nField = "d"
            Else
                nStart = 6
                nLength = 4
                nField = "yyyy"
            End If
            
            If scrMain(Index).Value < 0 Then
                txtMain(Index).Text = Format(DateAdd(nField, 1, CDate(Format(txtMain(Index).Text, "mm/dd/yyyy"))), "mm/dd/yyyy")
            Else
                txtMain(Index).Text = Format(DateAdd(nField, -1, CDate(Format(txtMain(Index).Text, "mm/dd/yyyy"))), "mm/dd/yyyy")
            End If
            
        Case dsTime
            If txtMain(Index).SelStart < 2 Or (txtMain(Index).SelStart = 2 And Left(Trim(txtMain(Index).Text), 1) = "0") Then
                nStart = 0
                nLength = 1
                nField = "h"
            ElseIf txtMain(Index).SelStart >= 2 And txtMain(Index).SelStart < 5 Then
                nStart = 2
                nLength = 2
                nField = "n"
            Else
                nStart = 5
                nLength = 2
                nField = "AMPM"
            End If
            
            If nField = "AMPM" Then
                If InStr(1, UCase(txtMain(Index).Text), "AM") <> 0 Then
                    txtMain(Index).Text = Replace(UCase(txtMain(Index).Text), "AM", "PM")
                ElseIf InStr(1, UCase(txtMain(Index).Text), "PM") <> 0 Then
                    txtMain(Index).Text = Replace(UCase(txtMain(Index).Text), "PM", "AM")
                ElseIf InStr(1, UCase(txtMain(Index).Text), "A") <> 0 Then
                    txtMain(Index).Text = Replace(UCase(txtMain(Index).Text), "A", "PM")
                ElseIf InStr(1, UCase(txtMain(Index).Text), "P") <> 0 Then
                    txtMain(Index).Text = Replace(UCase(txtMain(Index).Text), "P", "AM")
                End If
            Else
                If scrMain(Index).Value < 0 Then
                    txtMain(Index).Text = Format(DateAdd(nField, 1, CDate(Format(txtMain(Index).Text, "h:nn AMPM"))), "h:nn AMPM")
                Else
                    txtMain(Index).Text = Format(DateAdd(nField, -1, CDate(Format(txtMain(Index).Text, "h:nn AMPM"))), "h:nn AMPM")
                End If
            End If
            If InStr(1, txtMain(Index).Text, ":") = 3 Then
                If nStart = 0 Then
                    nLength = 2
                Else
                    nStart = nStart + 1
                End If
            End If
            
        Case dsYear
            nStart = 0
            If scrMain(Index).Value < 0 Then
                txtMain(Index).Text = CInt(txtMain(Index).Text) + 1
            Else
                txtMain(Index).Text = CInt(txtMain(Index).Text) - 1
            End If
            nLength = Len(Trim(txtMain(Index).Text))
        
        Case dsMonth
            nStart = 0
            If scrMain(Index).Value < 0 Then
                txtMain(Index).Text = Format(DateAdd("m", 1, CDate(Trim(txtMain(Index).Text) & " 01,2000")), "mmmm")
            Else
                txtMain(Index).Text = Format(DateAdd("m", -1, CDate(Trim(txtMain(Index).Text) & " 01,2000")), "mmmm")
            End If
            nLength = Len(Trim(txtMain(Index).Text))
    
        Case dsMonthYear
            nStart = 0
            
            
            'If InStr(1, Trim(txtMain(Index).Text), " ") <> 0 Then
'                sMonth = Trim(Mid(Trim(txtMain(Index).Text), 1, InStr(1, Trim(txtMain(Index).Text), " ") - 1))
'                sYear = Trim(Mid(Trim(txtMain(Index).Text), InStrRev(Trim(txtMain(Index).Text), " ") + 1))
                
                nStart = 0
'                If txtMain(Index).SelStart > Len(sMonth) Then
                If Index = 1 Then
                    sYear = Trim(txtMain(Index).Text)
                    If IsNumeric(sYear) = True Then
                        If scrMain(Index).Value < 0 Then
                            sYear = CInt(sYear) + 1
                        Else
                            sYear = CInt(sYear) - 1
                        End If
                        txtMain(Index).Text = sYear
                        nStart = 0
                        nLength = Len(sYear)
                    End If
                Else
                    sMonth = Trim(txtMain(Index).Text)
                    If IsDate(sMonth & " 01,2000") = True Then
                        If scrMain(Index).Value < 0 Then
                            sMonth = Format(DateAdd("m", 1, CDate(Trim(sMonth) & " 01,2000")), "mmmm")
                        Else
                            sMonth = Format(DateAdd("m", -1, CDate(Trim(sMonth) & " 01,2000")), "mmmm")
                        End If
                        txtMain(Index).Text = sMonth
                        nStart = 0
                        nLength = Len(sMonth)
                    Else
                        txtMain(Index).Text = Format(date, "mmmm")
                    End If
                End If
            'Else
            '    txtMain(Index).Text = Format(Date, "mmmm")
            'End If
            
            'If IsDate(txtMain(Index).Text) = False Or (Len(sYear) <> 2 And Len(sYear) <> 4) Then
            '    txtMain(Index).Text = Format(Date, "mmmm")
            'End If
            
    End Select
    
    scrMain(Index).Value = 0
    txtMain(Index).SetFocus
    txtMain(Index).SelStart = nStart
    txtMain(Index).SelLength = nLength

    If m_Locked = False Then
        RaiseEvent Click
    End If
End Sub

Private Sub txtMain_Change(Index As Integer)
    Select Case m_DateMode
        Case dsDate, dsTime
            If IsDate(Trim(txtMain(Index).Text)) = True Then
                scrMain(Index).Enabled = True
            Else
                scrMain(Index).Enabled = False
            End If
            
        Case dsYear
            If Len(Trim(txtMain(Index).Text)) <> 2 And Len(Trim(txtMain(Index).Text)) <> 4 And IsNumeric(Trim(txtMain(Index).Text)) Then
                scrMain(Index).Enabled = False
            Else
                scrMain(Index).Enabled = True
            End If
        
        Case dsMonth
            If IsDate(Trim(txtMain(Index).Text) & " 01,2000") = False Then
                scrMain(Index).Enabled = False
            Else
                scrMain(Index).Enabled = True
            End If
            
        Case dsMonthYear
            If Index = 0 Then
                If IsDate(Trim(txtMain(Index).Text) & " 01,2000") = False Then
                    scrMain(Index).Enabled = False
                Else
                    scrMain(Index).Enabled = True
                End If
            Else
                If Len(Trim(txtMain(Index).Text)) <> 2 And Len(Trim(txtMain(Index).Text)) <> 4 And IsNumeric(Trim(txtMain(Index).Text)) Then
                    scrMain(Index).Enabled = False
                Else
                    scrMain(Index).Enabled = True
                End If
            End If
    End Select
    
    If m_Locked = False Then
        RaiseEvent Change
    End If
End Sub

Private Sub UserControl_GotFocus()
    If m_DateMode = dsMonthYear Then
        txtMain(0).SetFocus
    Else
        txtMain(1).SetFocus
    End If
'    txtMain(1).SetFocus
End Sub

Private Sub UserControl_Resize()
    If m_DateMode = dsMonthYear Then
        scrMain(0).Left = 0
        txtMain(0).Left = scrMain(0).Left + scrMain(0).Width
        txtMain(0).Width = (UserControl.Width - (scrMain(0).Width * 2)) * 0.65
        txtMain(1).Left = txtMain(0).Left + txtMain(0).Width
        txtMain(1).Width = (UserControl.Width - (scrMain(0).Width * 2)) * 0.35
        scrMain(1).Left = UserControl.Width - scrMain(1).Width
    Else
        scrMain(1).Left = UserControl.Width - scrMain(1).Width
        txtMain(1).Width = UserControl.Width - scrMain(1).Width - 20
    End If
    UserControl.Height = txtMain(1).Height
End Sub

Public Property Get Alignment() As AlignmentConstants
    Alignment = txtMain(1).Alignment
End Property
Public Property Let Alignment(ByVal vNewValue As AlignmentConstants)
    txtMain(0).Alignment = vNewValue
    txtMain(1).Alignment = vNewValue
End Property

Public Property Get DateMode() As eDateScrollMode
    DateMode = m_DateMode
End Property
Public Property Let DateMode(ByVal vNewValue As eDateScrollMode)
    m_DateMode = vNewValue
    If vNewValue = dsMonthYear Then
'        txtMain(0).Font.Bold = True
'        txtMain(1).Font.Bold = True
        txtMain(0).Alignment = vbCenter
        txtMain(1).Alignment = vbCenter
        txtMain(0).Visible = True
       
        Call UserControl_Resize
        scrMain_Change 0
        scrMain_Change 1
        txtMain(0).Text = Format(date, "mmmm")
        txtMain(1).Text = Format(date, "yyyy")
    Else
'        txtMain(0).Font.Bold = False
'        txtMain(1).Font.Bold = False
        txtMain(0).Visible = False
      
        Call UserControl_Resize
        scrMain_Change 1
    End If
End Property

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
    Text = txtMain(1).Text
End Property
Public Property Let Text(ByVal vNewValue As String)
    txtMain(1).Text = vNewValue
    scrMain_Change 1
End Property

Public Property Get Month() As String
    Month = txtMain(0).Text
End Property
Public Property Let Month(ByVal vNewValue As String)
    txtMain(0).Text = vNewValue
    txtMain_Change 0
    txtMain(0).SelStart = 0
    txtMain(0).SelLength = Len(txtMain(0).Text)
    txtMain(0).TabIndex = 0
End Property

Public Sub MonthSetFocus()
    With txtMain(0)
        On Error Resume Next
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
        On Error GoTo 0
    End With
End Sub

Public Sub YearSetFocus()
    With txtMain(1)
        On Error Resume Next
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
        On Error GoTo 0
    End With
End Sub

Public Property Get Year() As String
    Year = txtMain(1).Text
End Property
Public Property Let Year(ByVal vNewValue As String)
    txtMain(1).Text = vNewValue
    txtMain_Change 1
    txtMain(1).SelStart = 0
    txtMain(1).SelLength = Len(txtMain(1).Text)
    txtMain(1).TabIndex = 0
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = txtMain(1).BackColor
End Property
Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    txtMain(0).BackColor = vNewValue
    txtMain(1).BackColor = vNewValue
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = txtMain(1).ForeColor
End Property
Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
    txtMain(0).ForeColor = vNewValue
    txtMain(1).ForeColor = vNewValue
End Property

Public Property Get SelStart() As Integer
    SelStart = txtMain(1).SelStart
End Property
Public Property Let SelStart(ByVal vNewValue As Integer)
    txtMain(1).SelStart = vNewValue
End Property

Public Property Get Locked() As Boolean
    Locked = m_Locked
End Property
Public Property Let Locked(ByVal vNewValue As Boolean)
    m_Locked = vNewValue
End Property

Public Property Get SelLength() As Integer
    SelLength = txtMain(1).SelLength
End Property
Public Property Let SelLength(ByVal vNewValue As Integer)
    txtMain(1).SelLength = vNewValue
End Property

'Private Sub UserControl_Show()
'    If m_DateMode = dsMonthYear Then
'        txtMain(0).SetFocus
'    Else
'        txtMain(1).SetFocus
'    End If
'End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Alignment", Alignment)
    Call PropBag.WriteProperty("Text", Text)
    Call PropBag.WriteProperty("BackColor", BackColor)
    Call PropBag.WriteProperty("ForeColor", ForeColor)
    Call PropBag.WriteProperty("DateMode", DateMode)
    Call PropBag.WriteProperty("Month", Month)
    Call PropBag.WriteProperty("Year", Year)
End Sub
 Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Alignment = PropBag.ReadProperty("Alignment", Alignment)
    Text = PropBag.ReadProperty("Text", Text)
    BackColor = PropBag.ReadProperty("BackColor", BackColor)
    ForeColor = PropBag.ReadProperty("ForeColor", ForeColor)
    DateMode = PropBag.ReadProperty("DateMode", DateMode)
    Month = PropBag.ReadProperty("Month", Month)
    Year = PropBag.ReadProperty("Year", Year)
End Sub

