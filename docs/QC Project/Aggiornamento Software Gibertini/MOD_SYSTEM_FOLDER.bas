Attribute VB_Name = "MOD_SYSTEM_FOLDER"


Const CSIDL_DESKTOP = &H0
Const CSIDL_PROGRAMS = &H2
Const CSIDL_CONTROLS = &H3
Const CSIDL_PRINTERS = &H4
Const CSIDL_PERSONAL = &H5
Const CSIDL_FAVORITES = &H6
Const CSIDL_STARTUP = &H7
Const CSIDL_RECENT = &H8
Const CSIDL_SENDTO = &H9
Const CSIDL_BITBUCKET = &HA
Const CSIDL_STARTMENU = &HB
Const CSIDL_DESKTOPDIRECTORY = &H10
Const CSIDL_DRIVES = &H11
Const CSIDL_NETWORK = &H12
Const CSIDL_NETHOOD = &H13
Const CSIDL_FONTS = &H14
Const CSIDL_TEMPLATES = &H15
Const MAX_PATH = 260
Const CSIDL_APPDATA = &H1A

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long


Public PC_DOCUMENTI As String
Public USER_DOCUMENTI As String
Public USER_DESKTOP As String
Public APP_DATA_FOLDER As String
Public USER_DATA_PATH As String
Public USER_TEMP_PATH As String
Public USER_PATH As String
Public USER_PATH_RICHIESTE As String
Public LOG_PATH As String
Public MY_LOG_PATH As String
Public EurekaPath As String
Public EurekaPathPrev As String
Public USER_DATA_NETWORK As String
Public USER_UPDATE_PATH As String

'Dim fso As New FileSystemObject

Public Function Load_DOC_FOLDER()
    USER_DOCUMENTI = GetSpecialfolder(CSIDL_PERSONAL)
    PC_DOCUMENTI = USER_DOCUMENTI & "\"
End Function

Public Function Load_Folder(ByVal mSoftName As String) As Boolean
    Dim rc As Boolean
   ' Dim fso As New Scripting.FileSystemObject
    Dim fso As New FileSystemObject
    Dim Cartella As Folder
    Dim FileGenerico As File
     
     
   On Error GoTo ERR_FOLDER
   
rc = True

   ' Set fso = CreateObject("Scripting.FileSystemObject")


'-----------------------------------------------------------------------

    USER_DOCUMENTI = GetSpecialfolder(CSIDL_PERSONAL)
    PC_DOCUMENTI = USER_DOCUMENTI & "\"
    USER_DESKTOP = GetSpecialfolder(CSIDL_DESKTOP) & "\"
    APP_DATA_FOLDER = GetSpecialfolder(CSIDL_APPDATA)
  
'-----------------------------------------------------------------------

 USER_DOCUMENTI = USER_DOCUMENTI & "\Gibertini\" & mSoftName & "\"
 

    
    If fso.FolderExists(USER_DOCUMENTI & "update") = False Then
         fso.CreateFolder USER_DOCUMENTI & "update"
    End If
    
    
    USER_UPDATE_PATH = USER_DOCUMENTI & "update\"
    MsgBox USER_UPDATE_PATH
    
ERR_END:
    On Error GoTo 0
    'MsgBox USER_UPDATE_PATH
    Load_Folder = rc
    SaveSetting App.Title, "PATH", "folder vp4", rc
    Exit Function
ERR_FOLDER:
    rc = False
    MsgBox Err.Description
    Resume Next
End Function

Private Function GetSpecialfolder(CSIDL As Long) As String
    Dim r As Long
    Dim IDL As ITEMIDLIST
    'Get the special folder
    r = SHGetSpecialFolderLocation(200, CSIDL, IDL)
    If r = NOERROR Then
        'Create a buffer
        path$ = Space$(512)
        'Get the path from the IDList
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal path$)
        'Remove the unnecessary chr$(0)'s
        GetSpecialfolder = Left$(path, InStr(path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function






