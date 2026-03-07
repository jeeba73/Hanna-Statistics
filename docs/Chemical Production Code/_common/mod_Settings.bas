Attribute VB_Name = "mod_Settings"
Option Explicit

Public cIni As clsIniFile

Declare Function SendMessage& Lib "User" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, lParam As Any)

Global Const WM_USER = &H400
Global Const LB_DIR = (WM_USER + 14)

Global Const DIR_NORMALFILES = &H0
Global Const DIR_READONLY = &H8001
Global Const DIR_HIDDEN = &H8002
Global Const DIR_SYSTEM = &H8004
Global Const DIR_DIRECTORIES = &H8010
Global Const DIR_ARCHIVED = &H8020
Global Const DIR_DRIVES = &HC000


Public Specifiche(3) As String

Private bOpenFile As Boolean


Public Function OpenSettingDataFile(User_file_name As String, Optional Path As String)
    Dim INIpath As String
    Set cIni = New clsIniFile
    If Len(Path) = 0 Then
        Path = USER_PATH
    End If
    INIpath = Path & User_file_name

    
    cIni.InitFile INIpath, 1200 '1200 - UTF16-LE or 1251 (ANSI)
    cIni.CompareMethod = vbTextCompare
    
    bOpenFile = True

End Function

Public Function CloseSettingDataFile()
On Error Resume Next
If bOpenFile Then
     DoEvents
     cIni.Flush
    'when you finished work with the class
    Set cIni = Nothing
    bOpenFile = False
End If
   
End Function

Public Function SaveSettingData(User_file_name As String, PutKey As String, PutVariable As String, PutValue, Optional Path As String)

    
   ' If PutValue = "" Or IsNull(PutValue) Then Exit Function
    If User_file_name = "" Or IsNull(User_file_name) Then Exit Function
    
    If PutValue = "Falso" Then
        PutValue = False
    End If
    If PutValue = "Vero" Then
        PutValue = True
    End If
    
    
    If bOpenFile Then
        cIni.WriteParam PutKey, PutVariable, PutValue
    Else
        OpenSettingDataFile User_file_name, Path
        cIni.WriteParam PutKey, PutVariable, PutValue
    End If
    
   
    
End Function

Public Function GetSettingData(User_file_name As String, KEY As String, Variable As String, Optional DefValue, Optional Path As String)
Dim Temp As String

 
On Error GoTo ERR_READ:

    GetSettingData = ""

    If User_file_name = "" Or IsNull(User_file_name) Then Exit Function
    
    If bOpenFile Then
        Temp = cIni.ReadParam(KEY, Variable, DefValue)
    Else
        OpenSettingDataFile User_file_name, Path
        Temp = cIni.ReadParam(KEY, Variable, DefValue)
    End If
    
    If Temp = "Falso" Then Temp = False
    If Temp = "Vero" Then Temp = True
    
    GetSettingData = Temp
   ' MsgBox Path
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_READ:
   ' MsgBox err.Description
    Resume ERR_END:
    
End Function

