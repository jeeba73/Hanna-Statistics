VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAlarmGroup"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_Collection As Collection
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Function add(dateTime As Date, alarmType As eAlarmType, filename As String, message As String, Optional UID As Double, Optional repeat As Integer, Optional presetMins As Integer = 0, Optional alarmTime As Date) As Long
    Dim newOBJ As New cAlarm
    Dim nCount As Integer
    Dim nInsert As Integer
    
    nInsert = -1
    For nCount = 1 To m_Collection.Count
        If m_Collection(nCount).dateTime > dateTime Then
            nInsert = nCount
            Exit For
        End If
    Next nCount
    
    With newOBJ
        .dateTime = dateTime
        .alarmType = alarmType
        .filename = filename
        .message = message
        .repeat = repeat
        .presetMins = presetMins
        If CDbl(alarmTime) = 0 Then
            .alarmTime = DateAdd("n", -.presetMins, .dateTime)
        Else
            .alarmTime = alarmTime
        End If
        If UID = 0 Then
            .UID = GetTickCount + Int((1000 * Rnd) + 1) / 1000
        Else
            .UID = UID
        End If
    End With
    
    If nInsert = -1 Then
        m_Collection.add newOBJ
    Else
        m_Collection.add newOBJ, , nInsert
    End If
    add = newOBJ.UID
    Set newOBJ = Nothing
End Function
Public Sub DeleteUID(UID As Double)
    Dim nCount As Integer
    
    For nCount = 1 To m_Collection.Count
        If CStr(m_Collection(nCount).UID) = CStr(UID) Then
            m_Collection.Remove (nCount)
            Exit For
        End If
    Next nCount
End Sub
Public Sub Delete(Index As Integer)
    m_Collection.Remove Index
End Sub
Public Function Item(Index As Integer) As cAlarm
Attribute Item.VB_UserMemId = 0
    Set Item = m_Collection(Index)
End Function
Public Function Count() As Integer
    Count = m_Collection.Count
End Function
Private Sub Class_Initialize()
    Set m_Collection = New Collection
End Sub
Public Function FindItem(value As Double) As Integer
    Dim nCount As Integer
    
    For nCount = 1 To m_Collection.Count
        If CStr(m_Collection(nCount).UID) = CStr(value) Then
            FindItem = nCount
            Exit For
        End If
    Next nCount
End Function
Public Sub LoadData(filename As String)
    Dim nLoadFile As Integer
    Dim nextLine As String
    Dim sSplit() As String
    Dim sVersion As String
    
    Set m_Collection = New Collection
    If Dir(filename, vbNormal) <> "" Then
        nLoadFile = FreeFile
        Open filename For Input As #nLoadFile
            Line Input #nLoadFile, nextLine
            sVersion = nextLine
            Do While Not EOF(nLoadFile)
                Line Input #nLoadFile, nextLine
                If Trim(nextLine) <> "" Then
                    On Error GoTo e_Next
                    sSplit = Split(nextLine, SPACER_CHAR)
                    If CLng(sVersion) <= 1001004 Then
                        If UBound(sSplit) >= 4 Then
                            If CDate(sSplit(1)) >= DateAdd("n", -1, Now) Then
                                Call add(CDate(sSplit(1)), CInt(sSplit(2)), "", LoadFormat(sSplit(4)), CDbl(sSplit(0)), CInt(sSplit(3)))
                            End If
                        End If
                    ElseIf CLng(sVersion) <= 1001009 Then
                        If UBound(sSplit) >= 5 Then
                            If CDate(sSplit(1)) >= DateAdd("n", -1, Now) Then
                                Call add(CDate(sSplit(1)), CInt(sSplit(2)), CStr(sSplit(4)), LoadFormat(sSplit(5)), CDbl(sSplit(0)), CInt(sSplit(3)))
                            End If
                        End If
                    Else
                        If UBound(sSplit) >= 6 Then
                            If CDate(sSplit(1)) >= DateAdd("n", -1, Now) Then
                                Call add(CDate(sSplit(1)), CInt(sSplit(2)), CStr(sSplit(5)), LoadFormat(sSplit(6)), CDbl(sSplit(0)), CInt(sSplit(3)), CInt(sSplit(4)))
                            End If
                        End If
                    End If
e_Next:
                End If
            Loop
        Close #nLoadFile
    End If
End Sub
Public Sub SaveData(filename As String)
    Dim nLoadFile As Integer
    Dim nCount As Integer
    
    nLoadFile = FreeFile
    Open filename For Output As #nLoadFile
        Print #nLoadFile, Format(App.Major, "000") & Format(App.Minor, "000") & Format(App.Revision, "0000")
        For nCount = 1 To m_Collection.Count
            Print #nLoadFile, m_Collection(nCount).UID & SPACER_CHAR & m_Collection(nCount).dateTime & SPACER_CHAR & m_Collection(nCount).alarmType & SPACER_CHAR & m_Collection(nCount).repeat & SPACER_CHAR & m_Collection(nCount).presetMins & SPACER_CHAR & m_Collection(nCount).filename & SPACER_CHAR & SaveFormat(m_Collection(nCount).message)
        Next nCount
        Print #nLoadFile, ""
    Close #nLoadFile
End Sub

Public Function GetAlarmDays(inputDate As Date) As cAlarmGroup
    Dim nCount As Integer
    Dim dStartDate As Date
    Dim dEndDate As Date
    
    dStartDate = DateAdd("s", -1, Format(inputDate, "m") & "/01/" & Format(inputDate, "yy"))
    dEndDate = DateAdd("m", 1, Format(inputDate, "m") & "/01/" & Format(inputDate, "yy"))
    
    Set GetAlarmDays = New cAlarmGroup
    For nCount = 1 To m_Collection.Count
        With m_Collection(nCount)
            If .dateTime > dStartDate And .dateTime < dEndDate Then
                GetAlarmDays.add .dateTime, .alarmType, .filename, .message, .UID, .repeat, .presetMins
            End If
        End With
    Next nCount
    
End Function
Public Function GetEvents(inputDate As Date, searchInterval As eUpdate, Optional weekStartsWith As VbDayOfWeek = vbSunday) As cAlarmGroup
    Dim nCount As Integer
    Dim dStartDate As Date
    Dim dEndDate As Date
    Dim dHoldDate As Date
    Dim nDayOfWeek As Integer
    
    If searchInterval = ccDaily Then
        dStartDate = DateAdd("s", -1, inputDate)
        dEndDate = DateAdd("d", 1, inputDate)
    ElseIf searchInterval = ccWeekly Then
        nDayOfWeek = Format(inputDate, "w")
        nDayOfWeek = (nDayOfWeek + 7 - weekStartsWith) Mod 7
        dHoldDate = Format(DateAdd("d", -nDayOfWeek, inputDate), "mm/dd/yy")
        
        dStartDate = DateAdd("s", -1, dHoldDate)
        dEndDate = DateAdd("d", 7, dHoldDate)
    ElseIf searchInterval = ccMonthly Then
        dStartDate = DateAdd("s", -1, Format(inputDate, "m") & "/01/" & Format(inputDate, "yy"))
        dEndDate = DateAdd("m", 1, Format(inputDate, "m") & "/01/" & Format(inputDate, "yy"))
    End If
    
    Set GetEvents = New cAlarmGroup
    For nCount = 1 To m_Collection.Count
        If (m_Collection(nCount).dateTime > dStartDate And m_Collection(nCount).dateTime < dEndDate) Or searchInterval = ccNever Then
            GetEvents.add m_Collection(nCount).dateTime, m_Collection(nCount).alarmType, m_Collection(nCount).filename, m_Collection(nCount).message, m_Collection(nCount).UID, m_Collection(nCount).repeat
        End If
    Next nCount
    
End Function


