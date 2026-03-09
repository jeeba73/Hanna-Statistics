VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'IniFile class by Alex Dragokas

'Mainly intended for caching data beetween read-write operations
'Supports UTF-16 LE ini-files format
'Provides wide range of methods
'Doesn't support reading / saving commentary in ini file

#Const UseHashtable = False 'true, if you are using HashTable class by The Trick, instead of Scripting.Dictionary
#Const UseStringBuilder = True 'true, if you are using StringBuilder class by VolteFace (Fork by Dragokas) to improve speed a bit

Private mb_Loaded       As Boolean
Private mb_UnsavedData  As Boolean
Private mb_IsUnicode    As Boolean
Private ms_IniFile      As String
Private ml_Codepage     As Long
Private mo_Sect         As Object
Private mb_CompareMode  As VbCompareMethod

Private lWow64Old As Long
Private sWinSysDir As String
Private sWinDir As String
Private inIDE As Boolean
Private bIsWin64 As Boolean

Private Enum VB_FILE_ACCESS_MODE
    FOR_READ = 1
    FOR_READ_WRITE = 2
    FOR_OVERWRITE_CREATE = 4
End Enum

Private Type SAFEARRAY
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
End Type




Private Declare Function GetSystemWindowsDirectory Lib "kernel32.dll" Alias "GetSystemWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long
Private Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (OldValue As Long) As Long
Private Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (ByVal OldValue As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60.dll" (Src As Any, Dst As Any) As Long
Private Declare Function GetMem2 Lib "msvbvm60.dll" (Src As Any, Dst As Any) As Long
Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long) As Long
Private Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryW" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As Long, lpSecurityAttributes As Any) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpFileName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function IsWow64Process Lib "kernel32.dll" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfByConstesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, lpFileSize As Any) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpStringDest As Long, ByVal lpStringSrc As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long

Private Const MAX_PATH                  As Long = 260&
Private Const MAX_FILE_SIZE             As Currency = 157286400@ '150 MB. limit for file size to open
Private Const INVALID_HANDLE_VALUE      As Long = &HFFFFFFFF
Private Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10&
Private Const INVALID_FILE_ATTRIBUTES   As Long = -1&
Private Const GENERIC_READ              As Long = &H80000000
Private Const FILE_SHARE_READ           As Long = &H1
Private Const FILE_SHARE_WRITE          As Long = &H2
Private Const OPEN_EXISTING             As Long = 3&
Private Const GENERIC_WRITE             As Long = &H40000000
Private Const FILE_SHARE_DELETE         As Long = 4&
Private Const CREATE_NEW                As Long = 1&
Private Const CREATE_ALWAYS             As Long = 2&
Private Const NO_ERROR                  As Long = 0&
Private Const FILE_BEGIN                As Long = 0&
Private Const FILE_CURRENT              As Long = 1&
Private Const INVALID_SET_FILE_POINTER  As Long = &HFFFFFFFF
Private Const FILE_END                  As Long = 2&
Private Const FILE_ATTRIBUTE_READONLY   As Long = 1&
Private Const ERROR_FILE_NOT_FOUND      As Long = 2&


'// To begin working with the class (supported codepages are 1200 (UTF16-LE) and 1251 (ANSI).
Public Sub InitFile(sFile As String, id_Codepage As Long)
    On Error GoTo ErrorHandler:
    Dim sIniFile As String
    sIniFile = sFile
    If Not FileExists(sFile) Then
        'relative path -> search on %windir%, %windir%\System32
        If InStr(sFile, "\") = 0 Then
            If FileExists(sWinSysDir & "\" & sFile) Then
                sIniFile = sWinSysDir & "\" & sFile
            ElseIf FileExists(sWinDir & "\" & sFile) Then
                sIniFile = sWinDir & "\" & sFile
            End If
        End If
    End If
    'if another file is already loaded
    If (ms_IniFile <> "") Then
        'if the same file
        If StrComp(ms_IniFile, sIniFile, 1) = 0 And ml_Codepage = id_Codepage Then
            'do nothing
        Else
        'if another file or codepage
            'flush data to the last opened file
            Flush
            'reinit dictionary
            ReleaseDictionary
            InitDictionary
            
            ms_IniFile = sIniFile
            InitCodepage id_Codepage
            LoadFile
        End If
    Else 'first init
        ms_IniFile = sIniFile
        InitCodepage id_Codepage
        LoadFile
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg err, "clsIniFile_InitFile", sFile, id_Codepage
    If inIDE Then Stop: Resume Next
End Sub

Private Sub InitCodepage(id_Codepage As Long)
    On Error GoTo ErrorHandler:
    If id_Codepage <> 1251 And id_Codepage <> 1200 Then
        MsgBoxW "Incorrect codepage!"
        Exit Sub
    End If
    ml_Codepage = id_Codepage
    mb_IsUnicode = IIf(ml_Codepage = 1200, True, False)
    Exit Sub
ErrorHandler:
    ErrorMsg err, "clsIniFile_InitCodepage", id_Codepage
    If inIDE Then Stop: Resume Next
End Sub

'// Method of comparing the strings
Public Property Get CompareMethod() As VbCompareMethod
    CompareMethod = mb_CompareMode
End Property

Public Property Let CompareMethod(ByVal CompareMode As VbCompareMethod)
    mb_CompareMode = CompareMode
End Property

Private Sub LoadFile()
    On Error GoTo ErrorHandler:

    Dim aContents() As String
    Dim i As Long, j As Long, pos As Long
    Dim sSection As String
    Dim sParam As String
    Dim sData As String
    
    If Not FileExists(ms_IniFile) Then Exit Sub
    
    aContents = ReadFileToArray(ms_IniFile, mb_IsUnicode)
    
    If IsArrDimmed(aContents) Then
        Do
            aContents(i) = Trim$(aContents(i))
            
            If Left$(aContents(i), 1) = "[" Then
                'begin new section
                sSection = Mid$(aContents(i), 2, Len(aContents(i)) - 2)
                
                If Not mo_Sect.Exists(sSection) Then
                    #If UseHashtable Then
                        mo_Sect.Add sSection, New clsTrickHashTable
                    #Else
                        mo_Sect.Add sSection, CreateObject("Scripting.Dictionary")
                    #End If
                    
                    With mo_Sect(sSection)
                        .CompareMode = mb_CompareMode
                        For j = i + 1 To UBound(aContents)
                            If Left$(aContents(j), 1) = "[" Then Exit For
                            pos = InStr(aContents(j), "=")
                            If pos <> 0 And Not (Left$(aContents(j), 1) = ";") Then
                                sParam = Trim$(Left$(aContents(j), pos - 1))
                                sData = Trim$(Mid$(aContents(j), pos + 1))
                                If Not .Exists(sParam) Then .Add sParam, sData
                            End If
                        Next
                        i = j - 1
                    End With
                End If
            End If
            i = i + 1
        Loop Until i > UBound(aContents)
    End If
    mb_Loaded = True
    Exit Sub
ErrorHandler:
    ErrorMsg err, "clsIniFile_LoadFile", ms_IniFile
    If inIDE Then Stop: Resume Next
End Sub

Private Sub Class_Initialize()
    Dim lr As Long
    
    'init required variables
    sWinDir = String$(MAX_PATH, 0)
    
    lr = GetSystemWindowsDirectory(StrPtr(sWinDir), MAX_PATH)
    If lr Then
        sWinDir = Left$(sWinDir, lr)
        sWinSysDir = sWinDir & "\system32"
    End If

    inIDE = (App.LogMode = 0)
    
    bIsWin64 = IsWow64()

    mb_CompareMode = vbTextCompare
    InitDictionary
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler:
    Flush
    ReleaseDictionary
    Exit Sub
ErrorHandler:
    ErrorMsg err, "clsIniFile_Class_Terminate"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub InitDictionary()
    On Error GoTo ErrorHandler:
    #If UseHashtable Then
        Set mo_Sect = New clsTrickHashTable
    #Else
        Set mo_Sect = CreateObject("Scripting.Dictionary")
    #End If
    mo_Sect.CompareMode = mb_CompareMode
    Exit Sub
ErrorHandler:
    ErrorMsg err, "clsIniFile_InitDictionary"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub ReleaseDictionary()
    On Error GoTo ErrorHandler:
    Dim oKey As Variant
    If Not (mo_Sect Is Nothing) Then
        For Each oKey In mo_Sect.Keys
            Set mo_Sect(CStr(oKey)) = Nothing
        Next
    End If
    Set mo_Sect = Nothing
    Exit Sub
ErrorHandler:
    ErrorMsg err, "clsIniFile_ReleaseDictionary"
    If inIDE Then Stop: Resume Next
End Sub

'// Update phisical ini file from the cache
Public Function Flush() As Boolean
    On Error GoTo ErrorHandler:

    #If UseStringBuilder Then
        Dim sb As clsStringBuilder
    #Else
        Dim s As String
    #End If
    Dim sDir As String, sIniFile As String
    Dim oKey As Variant
    Dim oKey2 As Variant
    
    If mb_UnsavedData Then
        
        sIniFile = ms_IniFile
    
        sDir = GetParentDir(sIniFile)
        
        If Not FolderExists(sDir) Then
            If Not MkDirW(sDir) Then
                MsgBoxW "Could not create folder '" & sDir & "'. Please verify that write access is allowed to this location.", vbCritical
                Exit Function
            End If
        End If
        
        If Not CheckAccessWrite(sIniFile) Then
            MsgBoxW "The value could not be written to the settings file '" & sIniFile & "'. Please verify that write access is allowed to that file."
            Exit Function
        End If
        
        Dim oDict As Object
        
        #If UseStringBuilder Then
            Set sb = New clsStringBuilder
        
            If Not (mo_Sect Is Nothing) Then
                For Each oKey In mo_Sect.Keys 'each section
                    sb.Append "[" & oKey & "]" & vbCrLf
                    Set oDict = mo_Sect(CStr(oKey))
                    If Not (oDict Is Nothing) Then
                        For Each oKey2 In oDict.Keys 'each parameter
                            sb.Append oKey2 & "=" & oDict(CStr(oKey2)) & vbCrLf
                        Next
                        Set oDict = Nothing
                    End If
                    sb.Append vbCrLf
                Next
            End If
            Flush = WriteDataToFile(ms_IniFile, sb.ToString, mb_IsUnicode, True)
            Set sb = Nothing
        #Else
            If Not (mo_Sect Is Nothing) Then
                For Each oKey In mo_Sect.Keys 'each section
                    s = s & "[" & oKey & "]" & vbCrLf
                    Set oDict = mo_Sect(CStr(oKey))
                    If Not (oDict Is Nothing) Then
                        For Each oKey2 In oDict.Keys 'each parameter
                            s = s & oKey2 & "=" & oDict(CStr(oKey2)) & vbCrLf
                        Next
                        Set oDict = Nothing
                    End If
                    s = s & vbCrLf
                Next
            End If
            Flush = WriteDataToFile(ms_IniFile, s, mb_IsUnicode, True)
        #End If
        
        mb_UnsavedData = False
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsIniFile_Flush"
    If inIDE Then Stop: Resume Next
End Function

'// Read perameter data by its name + section name, where it is stored
Public Function ReadParam(sSection As Variant, sParameter As Variant, Optional vDefault As Variant = vbNullString) As String
    On Error GoTo ErrorHandler:
    Dim oDict As Object
    If ExistSection(sSection) Then
        Set oDict = mo_Sect(CStr(sSection))
        If oDict.Exists(CStr(sParameter)) Then
            ReadParam = oDict(CStr(sParameter))
        Else
            ReadParam = vDefault
        End If
        Set oDict = Nothing
    Else
        ReadParam = vDefault
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsIniFile_ReadParam", sSection, sParameter, vDefault
    'MsgBox err.Description
   Resume Next ' If inIDE Then Stop: Resume Next
End Function

'// Write (or replace) parameter data
Public Sub WriteParam(sSection As Variant, sParameter As Variant, vData As Variant)
    On Error GoTo ErrorHandler:
    Dim oDict As Object
    If ExistSection(sSection) Then
        Set oDict = mo_Sect(CStr(sSection))
        If oDict.Exists(CStr(sParameter)) Then
            If oDict(CStr(sParameter)) = vData Then
                Set oDict = Nothing
                Exit Sub
            Else
                oDict(CStr(sParameter)) = vData
            End If
        Else
            oDict.Add CStr(sParameter), vData
        End If
        Set oDict = Nothing
    Else
        #If UseHashtable Then
            mo_Sect.Add CStr(sSection), New clsTrickHashTable
        #Else
            mo_Sect.Add CStr(sSection), CreateObject("Scripting.Dictionary")
        #End If
        mo_Sect(CStr(sSection)).CompareMode = mb_CompareMode
        mo_Sect(CStr(sSection)).Add CStr(sParameter), vData
    End If
    mb_UnsavedData = True
    Exit Sub
ErrorHandler:
    ErrorMsg err, "clsIniFile_ReadParam", sSection, sParameter, vData
    If inIDE Then Stop: Resume Next
End Sub

'// Get the list of all sections to array
Public Function GetSections() As Variant()
    On Error GoTo ErrorHandler:
    Dim bExist As Boolean
    If Not (mo_Sect Is Nothing) Then
        If mo_Sect.Count <> 0 Then
            GetSections = mo_Sect.Keys
            bExist = True
        End If
    End If
    If Not bExist Then GetSections = Array()
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsIniFile_GetSections"
    If inIDE Then Stop: Resume Next
End Function

'// Get the list of all parameter names to array by section name, where there are stored
Public Function GetParamNames(sSection As Variant) As Variant()
    On Error GoTo ErrorHandler:
    Dim oDict As Object
    Dim bExist As Boolean
    If ExistSection(sSection) Then
        Set oDict = mo_Sect(CStr(sSection))
        If oDict.Count <> 0 Then
            GetParamNames = oDict.Keys
            bExist = True
        End If
        Set oDict = Nothing
    End If
    If Not bExist Then GetParamNames = Array()
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsIniFile_GetParamNames", sSection
    If inIDE Then Stop: Resume Next
End Function

'// Get the list of all data of parameters to array by section name, where there are stored
Public Function GetParamValues(sSection As Variant) As Variant()
    On Error GoTo ErrorHandler:
    Dim oDict As Object
    Dim bExist As Boolean
    If ExistSection(sSection) Then
        Set oDict = mo_Sect(CStr(sSection))
        If oDict.Count <> 0 Then
            GetParamValues = oDict.Items
            bExist = True
        End If
        Set oDict = Nothing
    End If
    If Not bExist Then GetParamValues = Array()
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsIniFile_GetParamValues", sSection
    If inIDE Then Stop: Resume Next
End Function

'// Check if section name is exist
Public Function ExistSection(sSection As Variant) As Boolean
    On Error GoTo ErrorHandler:
    If Not (mo_Sect Is Nothing) Then
        ExistSection = mo_Sect.Exists(CStr(sSection))
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsIniFile_ExistSection", sSection
    If inIDE Then Stop: Resume Next
End Function

'// Check if parameter name in particular section is exist
Public Function ExistParam(sSection As Variant, sParameter As Variant) As Boolean
    On Error GoTo ErrorHandler:
    If ExistSection(sSection) Then
        ExistParam = mo_Sect(CStr(sSection)).Exists(CStr(sParameter))
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsIniFile_ExistParam", sSection, sParameter
    If inIDE Then Stop: Resume Next
End Function

'// Check if concrete data is exist in section
Public Function ExistData(sSection As Variant, sData As Variant) As Boolean
    On Error GoTo ErrorHandler:
    Dim aNames() As Variant
    Dim i As Long
    aNames = GetParamNames(sSection)
    If IsArrDimmed(aNames) Then
        For i = 0 To UBound(aNames)
            If StrComp(CStr(sData), ReadParam(sSection, aNames(i)), mb_CompareMode) = 0 Then
                ExistData = True
                Exit For
            End If
        Next
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsIniFile_ExistData", sSection, sData
    If inIDE Then Stop: Resume Next
End Function

'// Get first parameter name by data stored there in particular section
Public Function GetParamNameByData(sSection As Variant, sData As Variant) As String
    On Error GoTo ErrorHandler:
    Dim aNames() As Variant
    Dim i As Long
    aNames = GetParamNames(sSection)
    If IsArrDimmed(aNames) Then
        For i = 0 To UBound(aNames)
            If StrComp(CStr(sData), ReadParam(sSection, aNames(i)), mb_CompareMode) = 0 Then
                GetParamNameByData = aNames(i)
                Exit For
            End If
        Next
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsIniFile_GetParamNameByData", sSection, sData
    If inIDE Then Stop: Resume Next
End Function

'// Create empty section, if it is not exist yet
Public Function CreateSection(sSection As Variant) As Boolean
    On Error GoTo ErrorHandler:
    If Not ExistSection(sSection) Then
        #If UseHashtable Then
            mo_Sect.Add CStr(sSection), New clsTrickHashTable
        #Else
            mo_Sect.Add CStr(sSection), CreateObject("Scripting.Dictionary")
        #End If
        mo_Sect(CStr(sSection)).CompareMode = mb_CompareMode
        CreateSection = True
        mb_UnsavedData = True
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsIniFile_CreateSection", sSection
    If inIDE Then Stop: Resume Next
End Function

'// Remove ALL sections
Public Sub RemoveSectionsAll()
    On Error GoTo ErrorHandler:
    ReleaseDictionary
    InitDictionary
    Exit Sub
ErrorHandler:
    ErrorMsg err, "clsIniFile_RemoveSectionsAll"
    If inIDE Then Stop: Resume Next
End Sub

'// Remove concrete section with all parameters in it
Public Function RemoveSection(sSection As Variant) As Boolean
    On Error GoTo ErrorHandler:
    If ExistSection(sSection) Then
        mo_Sect.Remove CStr(sSection)
        RemoveSection = True
        mb_UnsavedData = True
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsIniFile_RemoveSection", sSection
    If inIDE Then Stop: Resume Next
End Function

'// Remove parameter in particular section
Public Function RemoveParam(sSection As Variant, sParameter As Variant) As Boolean
    On Error GoTo ErrorHandler:
    If ExistParam(sSection, sParameter) Then
        mo_Sect(CStr(sSection)).Remove CStr(sParameter)
        RemoveParam = True
        mb_UnsavedData = True
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsIniFile_RemoveParam", sSection, sParameter
    If inIDE Then Stop: Resume Next
End Function

'// Get currently loaded file name
Public Property Get Filename() As String
    Filename = ms_IniFile
End Property

'// Get number of sections
Public Property Get CountSections() As Long
    On Error GoTo ErrorHandler:
    If Not (mo_Sect Is Nothing) Then
        CountSections = mo_Sect.Count
    End If
    Exit Property
ErrorHandler:
    ErrorMsg err, "clsIniFile_CountSections"
    If inIDE Then Stop: Resume Next
End Property

'// Get number of parameters within section
Public Property Get CountParams(sSection As Variant) As Long
    On Error GoTo ErrorHandler:
    If Not (mo_Sect Is Nothing) Then
        If ExistSection(sSection) Then
            CountParams = mo_Sect(CStr(sSection)).Count
        End If
    End If
    Exit Property
ErrorHandler:
    ErrorMsg err, "clsIniFile_CountParams", sSection
    If inIDE Then Stop: Resume Next
End Property

'###############################
'
' Helper functions
'
'###############################

Private Function FileExists(ByVal sFile As String, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    
    Dim Redirect As Boolean
    Dim bOldStatus As Boolean
    Dim ret As Long
    
    If Not bUseWow64 Then Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    
    ret = GetFileAttributes(StrPtr(sFile))
    If ret <> INVALID_HANDLE_VALUE And (0 = (ret And FILE_ATTRIBUTE_DIRECTORY)) Then
        FileExists = True
    ElseIf err.LastDllError = 5 Then
        FileExists = True
    End If
    
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "modFile.FileExists", "File:", sFile
    If inIDE Then Stop: Resume Next
End Function

Private Function FolderExists(ByVal sFolder$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:

    Dim ret As Long, Redirect As Boolean, bOldStatus As Boolean
    
    sFolder = Trim$(sFolder)
    If Len(sFolder) = 0 Then Exit Function
    If Left$(sFolder, 2) = "\\" Then Exit Function 'network path
    
    If Not bUseWow64 Then Redirect = ToggleWow64FSRedirection(False, sFolder, bOldStatus)
    
    ret = GetFileAttributes(StrPtr(sFolder))
    If CBool(ret And vbDirectory) And (ret <> INVALID_FILE_ATTRIBUTES) Then
        FolderExists = True
    ElseIf err.LastDllError = 5 Then
        FolderExists = True
    End If
    
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "modFile.FolderExists", "Folder:", sFolder$, "Redirection: ", bUseWow64
    If inIDE Then Stop: Resume Next
End Function

Private Function ToggleWow64FSRedirection(bEnable As Boolean, Optional PathNecessity As String, Optional OldStatus As Boolean) As Boolean
    'Static lWow64Old        As Long    'Warning: do not use initialized variables for this API !
                                        'Static variables is not allowed !
                                        'lWow64Old is now declared globally
    'True - enable redirector
    'False - disable redirector

    'OldStatus: current state of redirection
    'True - redirector was enabled
    'False - redirector was disabled

    'Return value is:
    'true if success

    Static IsNotRedirected  As Boolean
    Dim lr                  As Long
    
    OldStatus = Not IsNotRedirected
    
    If Not bIsWin64 Then Exit Function
    
    If Len(PathNecessity) <> 0 Then
        If StrComp(Left$(Replace(Replace(PathNecessity, "/", "\"), "\\", "\"), Len(sWinSysDir)), sWinSysDir, vbTextCompare) <> 0 Then Exit Function
    End If
    
    If bEnable Then
        If IsNotRedirected Then
            lr = Wow64RevertWow64FsRedirection(lWow64Old)
            ToggleWow64FSRedirection = (lr <> 0)
            IsNotRedirected = False
        End If
    Else
        If Not IsNotRedirected Then
            lr = Wow64DisableWow64FsRedirection(lWow64Old)
            ToggleWow64FSRedirection = (lr <> 0)
            IsNotRedirected = True
        End If
    End If
    
End Function

'stub
Private Function MsgBoxW(Prompt As String, Optional Buttons As VbMsgBoxStyle, Optional Title As String = " ") As VbMsgBoxResult
    MsgBoxW = MsgBox(Prompt, Buttons, Title)
End Function

Private Sub ErrorMsg(ParamArray aLog())
    Dim i As Long
    Dim s As String
    For i = 0 To UBound(aLog)
        s = s & aLog(i) & ", "
    Next
    Debug.Print Left$(s, Len(s) - 2)
End Sub

Private Function ReadFileToArray(sFile As String, Optional isUnicode As Boolean) As String()
    On Error GoTo ErrorHandler:
    Dim Text  As String
    Text = Replace$(ReadFileContents(sFile, isUnicode), vbCr, vbNullString)
    If Len(Text) <> 0 Then
        ReadFileToArray = Split(Text, vbLf)
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "ReadFileToArray"
    If inIDE Then Stop: Resume Next
End Function

Private Function ReadFileContents(sFile As String, isUnicode As Boolean) As String
    On Error GoTo ErrorHandler:
    Dim hFile   As Long
    Dim b()     As Byte
    Dim Text    As String
    Dim lSize   As Currency
    Dim Redirect As Boolean, bOldStatus As Boolean
    If Not FileExists(sFile) Then Exit Function
    Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    OpenW sFile, FOR_READ, hFile
    If hFile = 0 Then Exit Function
    lSize = LOFW(hFile)
    If lSize = 0 Then CloseW hFile: Exit Function
    ReDim b(lSize - 1)
    GetW hFile, 1, , VarPtr(b(0)), UBound(b) + 1
    CloseW hFile
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If isUnicode Then
        ReadFileContents = b()
        If b(0) = &HFF& And b(1) = &HFE& Then ReadFileContents = Mid$(ReadFileContents, 2)  ' - BOM UTF16-LE
    Else
        ReadFileContents = StrConv(b(), vbUnicode, &H419&)
        If b(0) = &HEF& And b(1) = &HBB& And b(2) = &HBF& Then      ' - BOM UTF-8
            ReadFileContents = Mid$(ReadFileContents, 4)
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "ReadFileContents"
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function

Private Function IsArrDimmed(vArray As Variant) As Boolean
    IsArrDimmed = (GetArrDims(vArray) > 0)
End Function

Private Function GetArrDims(vArray As Variant) As Integer
    Dim ppSA As Long
    Dim pSA As Long
    Dim vt As Long
    Dim sa As SAFEARRAY
    Const vbByRef As Integer = 16384

    If IsArray(vArray) Then
        GetMem4 ByVal VarPtr(vArray) + 8, ppSA      ' pV -> ppSA (pSA)
        If ppSA <> 0 Then
            GetMem2 vArray, vt
            If vt And vbByRef Then
                GetMem4 ByVal ppSA, pSA                 ' ppSA -> pSA
            Else
                pSA = ppSA
            End If
            If pSA <> 0 Then
                memcpy sa, ByVal pSA, LenB(sa)
                If sa.pvData <> 0 Then
                    GetArrDims = sa.cDims
                End If
            End If
        End If
    End If
End Function

Private Function GetParentDir(sPath As String) As String
    Dim pos As Long
    pos = InStrRev(sPath, "\")
    If pos <> 0 Then
        GetParentDir = Left$(sPath, pos - 1)
    End If
End Function

Private Function MkDirW(ByVal Path As String, Optional ByVal LastComponentIsFile As Boolean = False) As Boolean
    ' Create folders struct
    ' LastComponentIsFile - true, if you specify filename as a last part of path component
    ' Return value: true, if successfully created or if folder is already exists
    Dim FC As String, lr As Boolean, pos As Long
    Dim bRedirect As Boolean, bOldStatus As Boolean
    If LastComponentIsFile Then Path = Left(Path, InStrRev(Path, "\") - 1) ' cut off file name
    If InStr(Path, ":") = 0 Then 'if replative path
        Dim sCurDir$, nChar As Long
        sCurDir = String$(MAX_PATH, 0&)
        nChar = GetCurrentDirectory(MAX_PATH + 1, StrPtr(sCurDir))
        sCurDir = Left$(sCurDir, nChar)
        If Right$(sCurDir, 1) <> "\" Then sCurDir = sCurDir & "\"
        Path = sCurDir & Path
    End If
    bRedirect = ToggleWow64FSRedirection(False, Path, bOldStatus)
    Do 'looping through each path component
        pos = pos + 1
        pos = InStr(pos, Path, "\")
        If pos Then FC = Left(Path, pos - 1) Else FC = Path
        If FolderExists(FC) Then
            lr = True 'if folder is already created
        Else
            lr = CBool(CreateDirectory(StrPtr(FC), ByVal 0&))
            If lr = 0 Then Exit Do
        End If
    Loop While (pos <> 0) And (lr <> 0)
    MkDirW = lr
    If bRedirect Then ToggleWow64FSRedirection bOldStatus
End Function

Private Function CheckAccessWrite(sFilePath As String, Optional bDeleteFile As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hFile As Long
    Dim bRedirect As Boolean
    Dim bOldStatus As Boolean
    
    bRedirect = ToggleWow64FSRedirection(False, sFilePath, bOldStatus)
    
    If FileExists(sFilePath) Then
        hFile = CreateFile(StrPtr(sFilePath), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    Else
        hFile = CreateFile(StrPtr(sFilePath), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, CREATE_NEW, ByVal 0&, ByVal 0&)
    End If
    
    If hFile > 0 Then
        CloseHandle hFile
        CheckAccessWrite = True
    End If
        
    If bRedirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    If bDeleteFile Then
        If FileExists(sFilePath) Then
            DeleteFileWEx StrPtr(sFilePath)
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "CheckAccessWrite"
    If inIDE Then Stop: Resume Next
End Function

Private Function WriteDataToFile(sFile As String, sContents As String, Optional isUnicode As Boolean, Optional bShowWarning As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hFile As Long
    Dim Redirect As Boolean, bOldStatus As Boolean
    Dim iAttr As Long
    
    Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    
    iAttr = GetFileAttributes(StrPtr(sFile))
    If (iAttr And 2048) Then iAttr = iAttr And Not 2048
    
    If 0 = DeleteFileWEx(StrPtr(sFile)) Then
        If bShowWarning Then
            MsgBoxW "The value '" & sContents & "' could not be written to the settings file '" & sFile & "'. Please verify that write access is allowed to that file.", vbCritical
        End If
        Exit Function
    End If
    
    OpenW sFile, FOR_OVERWRITE_CREATE, hFile
    If hFile <> 0 Then
        If isUnicode Then
            PutW hFile, 1, StrPtr(ChrW$(-257)), 2
            WriteDataToFile = PutW(hFile, 3, StrPtr(sContents), LenB(sContents), True)
        Else
            WriteDataToFile = PutW(hFile, 1, StrPtr(StrConv(sContents, vbFromUnicode)), Len(sContents))
        End If
        CloseW hFile
    End If
    
    If iAttr <> 0 Then
        SetFileAttributes StrPtr(sFile), iAttr
    End If
    
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    Exit Function
ErrorHandler:
    ErrorMsg err, "WriteDataToFile", "sFile=", sFile, "sContents=", sContents, "iAttr=", iAttr, "bShowWarning=", bShowWarning, "isUnicode=", isUnicode
    CloseW hFile
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function

Private Function IsWow64() As Boolean
    Dim hModule As Long, procAddr As Long, lIsWin64 As Long
    Static IsInit As Boolean, Result As Boolean

    If IsInit Then
        IsWow64 = Result
    Else
        IsInit = True
        hModule = LoadLibrary(StrPtr("kernel32.dll"))
        If hModule Then
            procAddr = GetProcAddress(hModule, "IsWow64Process")
            If procAddr <> 0 Then
                IsWow64Process GetCurrentProcess(), lIsWin64
                Result = CBool(lIsWin64)
                IsWow64 = Result
            End If
            FreeLibrary hModule
        End If
    End If
End Function

Private Function OpenW(Filename As String, Access As VB_FILE_ACCESS_MODE, retHandle As Long, Optional MountToMemory As Boolean) As Boolean '// TODO: MountToMemory
    
    Dim FSize As Currency
    
    If Access And (FOR_READ Or FOR_READ_WRITE) Then
        If Not FileExists(Filename) Then
            retHandle = INVALID_HANDLE_VALUE
            Exit Function
        End If
    End If
        
    If Access = FOR_READ Then
        retHandle = CreateFile(StrPtr(Filename), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    ElseIf Access = FOR_OVERWRITE_CREATE Then
        retHandle = CreateFile(StrPtr(Filename), GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, CREATE_ALWAYS, ByVal 0&, ByVal 0&)
    ElseIf Access = FOR_READ_WRITE Then
        retHandle = CreateFile(StrPtr(Filename), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    Else
        'WriteCon "Wrong access mode!", cErr
    End If

    OpenW = (INVALID_HANDLE_VALUE <> retHandle)
    
    ' ограничение на максимально возможный файл для открытия ( > 100 МБ )
    If OpenW Then
        If Access And (FOR_READ Or FOR_READ_WRITE) Then
            FSize = LOFW(retHandle)
            If FSize > MAX_FILE_SIZE Then
                CloseHandle retHandle
                retHandle = INVALID_HANDLE_VALUE
                OpenW = False
                '"Не хочу и не буду открывать этот файл, потому что его размер превышает безопасный максимум"
                err.Clear: Debug.Print err, "modFile.OpenW", "Trying to open too big file" & ": (" & (FSize \ 1024 \ 1024) & " MB.) " & Filename
            End If
        End If
    Else
        Debug.Print err, "modFile.OpenW", "Cannot open file: " & Filename
        'Err.Raise 75 ' Path/File Access error
    End If

End Function

                                                                  'do not change Variant type at all or you will die ^_^
Private Function GetW(hFile As Long, pos As Long, Optional vOut As Variant, Optional vOutPtr As Long, Optional cbToRead As Long) As Boolean
    
    'On Error GoTo ErrorHandler
    
    Dim lBytesRead  As Long
    Dim lr          As Long
    Dim ptr         As Long
    Dim vType       As Long
    Dim UnknType    As Boolean
    Dim oldPos      As Long
    
    If pos >= 1 Then
        pos = pos - 1   ' VB's Get & SetFilePointer difference correction
        
        oldPos = SetFilePointer(hFile, 0&, ByVal 0&, FILE_CURRENT)
        If oldPos = INVALID_SET_FILE_POINTER And NO_ERROR <> err.LastDllError Then
            err.Clear: ErrorMsg err, "Cannot get file pointer! LastDllErr = " & err.LastDllError: err.Raise 52
        End If
        
        If pos <> oldPos Then
            If INVALID_SET_FILE_POINTER = SetFilePointer(hFile, pos, ByVal 0&, FILE_BEGIN) Then
                If NO_ERROR <> err.LastDllError Then
                    err.Clear: ErrorMsg err, "Cannot set file pointer! LastDllErr = " & err.LastDllError: err.Raise 52
                End If
            End If
        End If
    End If
    
    If Not IsMissing(vOut) Then
        vType = VarType(vOut)
    End If
    
    If 0 <> cbToRead Then   'vbError = vType
        lr = ReadFile(hFile, vOutPtr, cbToRead, lBytesRead, 0&)
                
    ElseIf vbString = vType Then
        lr = ReadFile(hFile, StrPtr(vOut), Len(vOut), lBytesRead, 0&)
        If err.LastDllError <> 0 Or lr = 0 Then err.Raise 52
        
        vOut = StrConv(vOut, vbUnicode)
        If Len(vOut) <> 0 Then vOut = Left$(vOut, Len(vOut) \ 2)
    Else
        'do a bit of magik :)
        memcpy ptr, ByVal VarPtr(vOut) + 8, 4& 'VT_BYREF
        Select Case vType
        Case vbByte
            lr = ReadFile(hFile, ptr, 1&, lBytesRead, 0&)
        Case vbInteger
            lr = ReadFile(hFile, ptr, 2&, lBytesRead, 0&)
        Case vbLong
            lr = ReadFile(hFile, ptr, 4&, lBytesRead, 0&)
        Case vbCurrency
            lr = ReadFile(hFile, ptr, 8&, lBytesRead, 0&)
        Case Else
            UnknType = True
            ErrorMsg err, "modFile.GetW. type #" & VarType(vOut) & " of buffer is not supported.": err.Raise 52
        End Select
    End If
    GetW = (0 <> lr)
    If 0 = lr And Not UnknType Then ErrorMsg err, "Cannot read file. LastDllErr = " & err.LastDllError: err.Raise 52
    
'    Exit Function
'ErrorHandler:
'    AppendErrorLogFormat Now, err, "modFile.GetW"
'    Resume Next
End Function

Private Function PutW(hFile As Long, pos As Long, vInPtr As Long, cbToWrite As Long, Optional doAppend As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim lBytesWrote  As Long
    
    pos = pos - 1   ' VB's Get & SetFilePointer difference correction
    
    If doAppend Then
        If INVALID_SET_FILE_POINTER = SetFilePointer(hFile, 0&, ByVal 0&, FILE_END) Then Exit Function
    Else
        If INVALID_SET_FILE_POINTER = SetFilePointer(hFile, pos, ByVal 0&, FILE_BEGIN) Then Exit Function
    End If
    
    If NO_ERROR = err.LastDllError Then
    
        If WriteFile(hFile, vInPtr, cbToWrite, lBytesWrote, 0&) Then PutW = True
        
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "modFile.PutW"
End Function

Private Function LOFW(hFile As Long) As Currency
    On Error GoTo ErrorHandler
    Dim lr          As Long
    Dim FileSize    As Currency
    
    If hFile Then
        lr = GetFileSizeEx(hFile, FileSize)
        If lr Then
            If FileSize < 10000000000@ Then
                LOFW = FileSize * 10000&
            Else
                err.Clear
                ErrorMsg Now, "File is too big. Size: " & FileSize
            End If
        End If
    End If
ErrorHandler:
End Function

Private Function CloseW(hFile As Long) As Long
    CloseW = CloseHandle(hFile)
End Function

Private Function DeleteFileWEx(lpSTR As Long, Optional ForceDeleteMicrosoft As Boolean, Optional DisallowRemoveOnReboot As Boolean) As Long
    On Error GoTo ErrorHandler:

    Dim iAttr As Long, lr As Long, sExt As String, sNewName As String
    Dim Redirect As Boolean, bOldStatus As Boolean

    Dim Filename$
    Filename = String$(lstrlen(lpSTR), vbNullChar)
    If Len(Filename) <> 0 Then
        lstrcpy StrPtr(Filename), lpSTR
    Else
        Exit Function
    End If
    
    Redirect = ToggleWow64FSRedirection(False, Filename, bOldStatus)
    
    iAttr = GetFileAttributes(lpSTR)
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    
    If iAttr And FILE_ATTRIBUTE_READONLY Then SetFileAttributes lpSTR, iAttr And Not FILE_ATTRIBUTE_READONLY
    lr = DeleteFileW(lpSTR)
    
    If lr <> 0 Then 'success
        DeleteFileWEx = lr
    End If
    
    If err.LastDllError = ERROR_FILE_NOT_FOUND Then
        DeleteFileWEx = 1
    End If
    
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "Parser.DeleteFileWEx", "File:", Filename
    If inIDE Then Stop: Resume Next
End Function
