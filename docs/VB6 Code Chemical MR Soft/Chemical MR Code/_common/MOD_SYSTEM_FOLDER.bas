Attribute VB_Name = "MOD_SYSTEM_FOLDER"
Public Const PROGRAM_NAME As String = "ChemicalMR"

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
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Public USER_DOCUMENTI As String
Public PC_DOCUMENTI As String
Public MY_LOG_PATH As String
Public LOG_PATH As String
Public USER_DESKTOP As String
Public APP_DATA_FOLDER As String


Public USER_DATA_PATH As String ' Recipe for STDPreparation
Public USER_TEMP_PATH As String ' Recipe for STDPreparation
Public USER_SCHEDULED_STD_PATH As String
Public USER_PREPARATION_PATH As String

Public USER_PATH As String
Public USER_UPDATE_PATH As String
Public USER_EXCEL_PATH As String
Public PathRequisition As String '
Public PathRecipe As String '
Public PathPictograms As String

Public PathReport As String '
Public USER_LABEL_PATH As String




 Public Function Load_Folder() As Boolean
    Dim FSO As New FileSystemObject
    Dim Cartella As Folder
    Dim a As Integer
    Dim FileGenerico As file
     
Dim rc As Boolean
rc = True
'-----------------------------------------------------------------------
    USER_DOCUMENTI = GetSpecialfolder(CSIDL_PERSONAL)
    PC_DOCUMENTI = USER_DOCUMENTI & "\"
    USER_DESKTOP = GetSpecialfolder(CSIDL_DESKTOP)
    APP_DATA_FOLDER = GetSpecialfolder(&H1A)
'-----------------------------------------------------------------------
a = 1

     
   On Error GoTo ERR_FOLDER

If Not FSO.FolderExists(USER_DOCUMENTI & "\Gibertini\") Then FSO.CreateFolder USER_DOCUMENTI & "\Gibertini\"
If Not FSO.FolderExists(USER_DOCUMENTI & "\Gibertini\" & PROGRAM_NAME & "\") Then FSO.CreateFolder USER_DOCUMENTI & "\Gibertini\" & PROGRAM_NAME & "\"

USER_DOCUMENTI = USER_DOCUMENTI & "\Gibertini\" & PROGRAM_NAME & "\"
    
a = 2
        

    If DirExists(USER_DOCUMENTI & "\bin") = False Then
    '-----------------------------------------------------------
    ' se non esiste la Dir BIN allora copio i files
    '-----------------------------------------------------------
        MakePath (USER_DOCUMENTI)
        CopyFolder App.PATH & "\bin", USER_DOCUMENTI & "\Bin"
        
    End If
   PathReport = "Report (" & Right$(year(Date), 2) & ")"
   
    If FSO.FolderExists(USER_DOCUMENTI & PathReport) = False Then
    '-----------------------------------------------------------
    ' se non esiste la Dir ATTESTATI ECC.. allora copio i files
    '-----------------------------------------------------------
            FSO.CreateFolder USER_DOCUMENTI & PathReport
    End If

    If FSO.FolderExists(USER_DOCUMENTI & "Pictograms") = False Then
            FSO.CreateFolder USER_DOCUMENTI & "Pictograms"
            FSO.CopyFolder App.PATH & "\Pictograms", USER_DOCUMENTI & "Pictograms"
    End If
             
             
    PathPictograms = USER_DOCUMENTI & "Pictograms"
     
    

a = 3
    
     If DirExists(USER_DOCUMENTI & "Temp") = False Then
        MakePath (USER_DOCUMENTI & "Temp")
        MakePath (USER_DOCUMENTI & "Data")
    End If

    If DirExists(USER_DOCUMENTI & "Excel") = False Then
        MakePath (USER_DOCUMENTI & "Excel")
    End If
    
     If DirExists(USER_DOCUMENTI & "Label") = False Then
        MakePath (USER_DOCUMENTI & "Label")
    End If
    
    USER_LABEL_PATH = USER_DOCUMENTI & "Label\"
    USER_EXCEL_PATH = USER_DOCUMENTI & "Excel"
    
    LOG_PATH = " Log file (" & Right$(year(Date), 2) & ")"
    
    If FSO.FolderExists(App.PATH & "\" & LOG_PATH) = False Then
        FSO.CreateFolder App.PATH & "\" & LOG_PATH
    End If
    
    MY_LOG_PATH = App.PATH & "\" & LOG_PATH
    
    
    USER_TEMP_PATH = USER_DOCUMENTI & "Temp\"
    USER_DATA_PATH = USER_DOCUMENTI & "Data\"
    
    
    'If DirExists(USER_DOCUMENTI & "ScheduledSTD") = False Then
        MakePath (USER_DOCUMENTI & "ScheduledSTD")
        MakePath (USER_DOCUMENTI & "Preparation")
        
        MakePath (USER_DOCUMENTI & "Preparation\Data")
        MakePath (USER_DOCUMENTI & "ScheduledSTD\Data")
    'End If
    
    

     USER_SCHEDULED_STD_PATH = USER_DOCUMENTI & "ScheduledSTD\"
     USER_PREPARATION_PATH = USER_TEMP_PATH
    
   
    USER_PATH = USER_TEMP_PATH
    
    
  a = 4
    
    
    APP_DATA_FOLDER = APP_DATA_FOLDER & "\Gibertini\" & PROGRAM_NAME & "\"
    
    If DirExists(APP_DATA_FOLDER) = False Then
    '----------------------------------------------------------
    ' se non esiste la Dir allora copio i files
    '----------------------------------------------------------
        
        MakePath (APP_DATA_FOLDER)
        FileCopy App.PATH & "\dBase\" & dbName, APP_DATA_FOLDER & dbName
        SaveSetting App.Title, "ARCHIVIO", "PATH", APP_DATA_FOLDER
        
    End If
    
    If FileExists(APP_DATA_FOLDER & dbName) Then
    Else
          FileCopy App.PATH & "\dBase\" & dbName, APP_DATA_FOLDER & dbName
          SaveSetting App.Title, "ARCHIVIO", "PATH", APP_DATA_FOLDER
    
    End If
    
    If FileExists(APP_DATA_FOLDER & dbCodeName) Then
    
    Else
          FileCopy App.PATH & "\dBase\" & dbCodeName, APP_DATA_FOLDER & dbCodeName
          SaveSetting App.Title, "ARCHIVIO", "PATH", APP_DATA_FOLDER
    
    End If



    If FSO.FolderExists(USER_DOCUMENTI & "update") = False Then
         FSO.CreateFolder USER_DOCUMENTI & "update"
    End If
a = 5
    
    USER_UPDATE_PATH = USER_DOCUMENTI & "update\"
    
ERR_END:
    On Error GoTo 0
    Load_Folder = rc
    SaveSetting App.Title, "PATH", "Folder ChemicalMR", rc
    Exit Function
ERR_FOLDER:
    rc = False
    MsgBox Err.Description & vbCrLf & "Err : " & a
    Resume Next



End Function
Private Function GetSpecialfolder(CSIDL As Long) As String
    Dim r As Long
    Dim IDL As ITEMIDLIST
    'Get the special folder
    r = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If r = NOERROR Then
        'Create a buffer
        PATH$ = Space$(512)
        'Get the path from the IDList
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal PATH$)
        'Remove the unnecessary chr$(0)'s
        GetSpecialfolder = Left$(PATH, InStr(PATH, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function






