Attribute VB_Name = "modGlobal"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SPACER_CHAR As String = "ÿ"

Public Function SaveFormat(inputString As String) As String
    SaveFormat = Replace(Replace(Replace(inputString, vbCrLf, Chr(1)), vbCr, Chr(2)), vbLf, Chr(3))
End Function
Public Function LoadFormat(inputString As String) As String
    LoadFormat = Replace(Replace(Replace(inputString, Chr(1), vbCrLf), Chr(2), vbCr), Chr(3), vbLf)
End Function
Public Function ShowFormat(inputString As String) As String
    ShowFormat = Replace(Replace(Replace(inputString, vbCrLf, " » "), vbCr, " » "), vbLf, " » ")
End Function
Public Function UnShowFormat(inputString As String) As String
    UnShowFormat = Replace(Replace(Replace(inputString, " » ", vbCrLf), " » ", vbCr), " » ", vbLf)
End Function

