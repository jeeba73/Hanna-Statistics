Attribute VB_Name = "mod_Week"
Option Explicit


Public Function Week(dteValue As Date) As Integer
   'Monday is set as first day of week
   Dim lngDate As Long
   Dim intWeek As Integer
   dteValue = FormatDateTime(dteValue, vbShortDate)
   'If january 1. is later then thursday, january 1. is not in week 1
   If Not weekday("01/01/" & year(dteValue), vbMonday) > 4 Then
      intWeek = 1
   Else
      intWeek = 0
   End If
   'Sets long-value for january 1.
   lngDate = CLng(CDate("01/01/" & year(dteValue)))
   
   'Finds the first monday of year
   lngDate = lngDate + (8 - weekday("01/01/" & year(dteValue), vbMonday))
   'Increases week by week until set date is passed
   While Not lngDate > CLng(CDate(dteValue))
      intWeek = intWeek + 1
      lngDate = lngDate + 7
   Wend
   'If the date set is not in week 1, this finds latest week previous year
   If intWeek = 0 Then
      intWeek = Week("31/12/" & year(dteValue) - 1)
   End If
   Week = intWeek
End Function

Public Function PreparationWeek(dteValue As Date) As String

Dim strWeek As String

    dteValue = FormatDateTime(dteValue, vbShortDate)
    strWeek = CStr(Week(dteValue))
    strWeek = strWeek & "/" & Right$(CStr(year(dteValue)), year(dteValue))
    PreparationWeek = strWeek
    
    
End Function

Public Function DateWeek(dteValue As String, ByRef MyDa As Date, ByRef MyA As Date) As Boolean

Dim var() As String
    DateWeek = True
    var = split(dteValue, "/")
    If UBound(var) = 0 Then
        ReDim Preserve var(1)
        var(1) = year(Now())
    End If
    
    If IsNumeric(var(0)) And IsNumeric(var(1)) Then
    
    MyDa = GetWeekStartDate(CInt(var(0)), CInt(var(1)))
    MyA = DateAdd("d", 7, MyDa)
   
    Else
        DateWeek = False
    End If
End Function

Private Function GetWeekStartDate(weekNumber As Integer, year As Integer) As Date

    Dim startDate As Date
    Dim day As Integer

    startDate = DateSerial(year, 1, 1)
    day = weekday(startDate, vbSunday)
    startDate = DateAdd("d", DaysToAdd(day), startDate)

    GetWeekStartDate = DateAdd("ww", weekNumber - 1, startDate)

End Function

Private Function DaysToAdd(day As Integer) As Integer

    DaysToAdd = 0
    If day > 1 Then DaysToAdd = 7 - day + 1

End Function
