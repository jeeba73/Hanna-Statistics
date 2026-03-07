Attribute VB_Name = "TaskManager"

Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, wParam As Any, lParam As Any) As Long
Public Const WM_CLOSE = &H10
Public Const WM_DESTROY = &H2

Public Declare Function SendMessageTimeout Lib "user32" _
    Alias "SendMessageTimeoutA" (ByVal hWnd As Long, _
    ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, _
    ByVal fuFlags As Long, ByVal uTimeout As Long, _
    pdwResult As Long) As Long
Public Const SMTO_BLOCK = &H1
Public Const SMTO_ABORTIFHUNG = &H2
Public Const WM_NULL = &H0
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
    Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Declare Function EnableWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public foreWORD
Public Function getalltopwindows(ByVal hWnd As Long, ByVal lParam As Long) As Long

Dim foregroundwindow As Long
Dim textlen As Long
Dim windowtext As String
Dim svar As Long
Static lastwindowtext As String



foregroundwindow = hWnd


textlen = GetWindowTextLength(foregroundwindow) + 1

windowtext = Space(textlen)
svar = GetWindowText(foregroundwindow, windowtext, textlen)
windowtext = Left(windowtext, Len(windowtext) - 1)

If windowtext = "" Then GoTo slask

'If Form1.Check2.Value = 1 Then
'If windowtext = Form1.Caption Then GoTo slask
processo = InStr(1, windowtext, App.EXEName)
If processo > 0 Then
foreWORD = foregroundwindow
'ora TERMINA tutti i WORD
svar = GetWindowThreadProcessId(foreWORD, nyprocessid)
procname = OpenProcess(PROCESS_ALL_ACCESS, 0&, nyprocessid)
svar2 = TerminateProcess(procname, 0&)
DoEvents

End If
'Form1.List1.ItemData(Form1.List1.NewIndex) = foregroundwindow
lastwindowtext = windowtext
slask:



getalltopwindows = 1
End Function

Public Function functionwindows(ByVal hWnd As Long, ByVal lParam As Long) As Long

Dim foregroundwindow As Long
Dim textlen As Long
Dim windowtext As String
Dim svar As Long
Static lastwindowtext As String



foregroundwindow = hWnd


textlen = GetWindowTextLength(foregroundwindow) + 1

windowtext = Space(textlen)
svar = GetWindowText(foregroundwindow, windowtext, textlen)
windowtext = Left(windowtext, Len(windowtext) - 1)

If windowtext = "" Then GoTo slask

'If Form1.Check2.Value = 1 Then
'If windowtext = Form1.Caption Then GoTo slask
processo = InStr(1, windowtext, "Taratura")
If processo > 0 Then
foreWORD = foregroundwindow
'ora TERMINA tutti i WORD
svar = GetWindowThreadProcessId(foreWORD, nyprocessid)
procname = OpenProcess(PROCESS_ALL_ACCESS, 0&, nyprocessid)
svar2 = TerminateProcess(procname, 0&)
DoEvents

End If
'Form1.List1.ItemData(Form1.List1.NewIndex) = foregroundwindow
lastwindowtext = windowtext
slask:



functionwindows = 1
End Function

