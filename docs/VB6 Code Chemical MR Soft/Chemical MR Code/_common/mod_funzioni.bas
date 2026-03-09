Attribute VB_Name = "mod_funzioni"
Option Explicit

Const FO_COPY = &H2
Const FO_DELETE = &H3
Const FO_MOVE = &H1
Const FO_RENAME = &H4
Const FOF_ALLOWUNDO = &H40
Const FOF_SILENT = &H4
Const FOF_NOCONFIRMATION = &H10
Const FOF_RENAMEONCOLLISION = &H8
Const FOF_NOCONFIRMMKDIR = &H200
Const FOF_FILESONLY = &H80

Private Type SHFILEOPSTRUCT
    hwnd      As Long
    wFunc     As Long
    pFrom     As String
    pTo       As String
    fFlags    As Integer
    fAborted  As Boolean
    hNameMaps As Long
    sProgress As String
End Type
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" _
    Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Public Declare Sub Sleep Lib "kernel32.dll" (ByVal milliSec As Long)




Public Function OpenWithDefault(ByVal FileName As String) As Boolean
 'ShellExecute returns a value greater than 32 if it was successful
    OpenWithDefault = (ShellExecute(0&, "", FileName, vbNullString, vbNullString, vbNormalFocus) > 32)
   ' MsgBox OpenWithDefault
End Function


Public Function GetDateCombo(ByVal Index As Integer) As String
Dim sString As String
Dim MyDate As Date
    Select Case Index
        Case 0
            MyDate = Date
            
        Case 1
            MyDate = Date - 30
        Case 2
            MyDate = Date - 365
        Case 3
            MyDate = Date - 36500
        Case Else
            MyDate = Date - 36500
           'Exit Function
    End Select
GetDateCombo = MyDate
End Function

Public Function FormatDataLAT(ByRef MyDataITA As String, Optional ByVal bHour As Boolean) As String
Dim NuovaData As String
On Error Resume Next
If IsNull(MyDataITA) Or MyDataITA = "" Then Exit Function
NuovaData = year(MyDataITA) & "/" & Format(Month(MyDataITA), "00") & "/" & Format(day(MyDataITA), "00") '& FormatDateTime(Now, vbShortTime)

If bHour Then
 NuovaData = NuovaData + " : " & FormatDateTime(Now(), vbShortTime)
End If
FormatDataLAT = NuovaData
End Function

Public Function FormatDataExp(ByRef MyDataITA As String) As String
Dim NuovaData As String
On Error Resume Next
NuovaData = Format(Month(MyDataITA), "00") & "/" & year(MyDataITA)  '& "/" & Format(Day(MyDataITA), "00") '& FormatDateTime(Now, vbShortTime)
FormatDataExp = NuovaData
End Function

Public Function FormatDataWeek(ByRef Week As String) As String
Dim NuovaData As String
On Error Resume Next
If InStr(Week, "/") Then
Else
    NuovaData = Format((Week), "00") & "/" & year(Now())  '& "/" & Format(Day(MyDataITA), "00") '& FormatDateTime(Now, vbShortTime)
End If
FormatDataWeek = NuovaData
End Function

Public Function SetExpDate(ByRef PreparationData As String, ByVal Exp As Double) As String
Dim Month As Double
Dim ExpData As String
On Error GoTo ERR_SET:

If IsDate(PreparationData) Then
    PreparationData = CDate(PreparationData)
    Select Case Int(Exp) = Exp
        
        Case True
            ExpData = DateAdd("yyyy", Exp, PreparationData)
        
        Case False
        
            Month = Int((Exp - Int(Exp)) * 12)
            ExpData = DateAdd("yyyy", Int(Exp), PreparationData)
            ExpData = DateAdd("m", Month, ExpData)
    End Select

    ExpData = FormatDataExp(ExpData)
    
Else




End If
SetExpDate = CStr(ExpData)

ERR_END:
    On Error GoTo 0
    Exit Function
ERR_SET:
    MsgBox Err.Description
    GoTo ERR_END:
End Function



Public Function FormatTimeLAT(ByRef MyDataITA As String) As String
Dim NuovaData As String
Dim Min As Double
On Error Resume Next
Min = Int(Minute(MyDataITA) / 5) * 5
NuovaData = Hour(MyDataITA) & ":" & Format$(Min, "00")
FormatTimeLAT = NuovaData
End Function


Public Function FormatDecimal(ByVal str As String) As String
Dim i As Integer
Dim Decimali As Integer
    FormatDecimal = "#0"
    If IsNumeric(str) Then
        Decimali = CInt(str)
    ElseIf str = "" Then
        FormatDecimal = "#0"
    Else
         FormatDecimal = "#0"
        Decimali = 0
        
    End If
    If Decimali > 0 Then
        FormatDecimal = FormatDecimal + "."
        Do
            i = i + 1
            FormatDecimal = FormatDecimal + "0"
        Loop Until i >= Decimali
    Else
        FormatDecimal = "#0"
    End If
        
End Function

Public Function GetLastImport(Optional ByRef L1 As String, Optional ByRef L2 As String)
Dim rc As Boolean
Dim str As String
Dim PATH As String
Dim Name As String
Dim strDate As String
    
    str = dbCodeName
    strDate = dbCodeRelease & " ( " & dbCodeDate & " - " & dbCodeOperator & ")"
    
   
  

       L1 = "File Name : " & str
       L2 = "Actual Rel: " & strDate
 
   
End Function


Public Function SetData(Periodo) As Date
    Dim DD As Date
    
    Select Case UCase(Periodo)
        Case UCase("Day")
            DD = Date
        Case UCase("Month")
            DD = Date - 31
        Case UCase("Year")
            DD = Date - 365
        Case Else
            DD = 0
    End Select
    SetData = DD
End Function




Public Function GetIndexArStr(AR() As String, ByVal Val As String) As Long
    Dim i As Long, ei As Long
    
    GetIndexArStr = -1
 
    On Error Resume Next
        ei = UBound(AR)
        If Err.NUMBER <> 0 Then Exit Function
     On Error GoTo 0
 
    For i = 0 To ei
        If AR(i, 0) = Val Then GetIndexArStr = i: Exit For
    Next
 
End Function


Public Function ResetUserDatabase() As Boolean
Dim rc As Boolean
rc = True
On Error GoTo ERR_RESET
If F_MsgBox.DoShow("Reset Database : All information will be lost.", "Reset Database", , "Reset", "No") Then
    
    dbChemicalMR.Close
    dbCode.Close
    
    FileCopy App.PATH & "\dBase\" & dbName, dbPath & dbName
    FileCopy App.PATH & "\dBase\" & dbCodeName, dbPath & dbCodeName
apri:
    If m_CreateArchivio(dbPath, MydbName) Then
            PopupMessage 2, ("Reset database : Operation done."), , , App.Title
            SaveSetting App.Title, "ImportExcel", "FileName", ""
            SaveSetting App.Title, "ImportExcel", "Date", ""
        Else
            rc = False
            PopupMessage 2, ("Warning : Database Error..."), , True, App.Title
            End
    End If
    
End If
ERR_END:
    If dbChemicalMR.State = 0 Then GoTo apri
    If dbCode.State = 0 Then GoTo apri
    ResetUserDatabase = rc
    Exit Function
ERR_RESET:
    rc = False
    PopupMessage 2, Err.Description
    Resume ERR_END:
End Function
