Attribute VB_Name = "mod_file"
Option Explicit
Private WordApp As Object


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
    hWnd      As Long
    wFunc     As Long
    pFrom     As String
    pTo       As String
    fFlags    As Integer
    fAborted  As Boolean
    hNameMaps As Long
    sProgress As String
End Type
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" _
    Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Public Declare Sub Sleep Lib "kernel32.dll" (ByVal milliSec As Long)
    
Public PROGRAM_NAME As String
Public PROGRAM_EXE_NAME As String
Public PROGRAM_VERSIONE As String
    
Public Function CopyFolder(txtSource As String, txtDestination As String) As Boolean
Dim lFileOp  As Long
Dim lresult  As Long
Dim lFlags   As Long
Dim SHFileOp As SHFILEOPSTRUCT
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    
            lFileOp = FO_COPY
    
    CopyFolder = True
    ' NOTE:  By adding the FOF_ALLOWUNDO flag you can move
    ' a file to the Recycle Bin instead of deleting it.
    '
    With SHFileOp
        .wFunc = lFileOp
        .pFrom = txtSource & vbNullChar & vbNullChar
        .pTo = txtDestination & vbNullChar & vbNullChar
      ' .fFlags
    End With
    lresult = SHFileOperation(SHFileOp)
    '
    ' If User hit Cancel button while operation is in progress,
    ' the fAborted parameter will be true
    '
    Screen.MousePointer = vbDefault
    
    If lresult <> 0 Or SHFileOp.fAborted Then
       ' MsgBox "Cartella Inesistente", vbInformation, "File Operations"
        CopyFolder = False
    End If
    
    On Error GoTo 0
End Function



Function FileExists(strFile As String) As Boolean
'********************************************************************************
'* Name : FileExists
'* Date : Feb-17, 2000
'* Author : David Costelloe
'* Returns : -1 = Does not exists 0 = Exists with zero bytes 1 = Exists > 0 Bytes
'*********************************************************************************
    Dim lSize As Long

    On Error Resume Next
    '* set lSize to -1
    lSize = -1
    'Get the length of the file
    lSize = FileLen(strFile)
    If lSize = 0 Then
        '* File is zero bytes and exists
        FileExists = True
    ElseIf lSize > 0 Then
        '* File Exists
        FileExists = True
    Else
        '* Does not exist
        FileExists = False
    End If
End Function




'Controlla se c'è già una istanza attiva
'se tipo = 0 default avvia una successiva istanza
'se tipo = 1 termina la nuova istanza con un messaggio
'se tipo = 2 passa il controllo alla prima
'N.B. bisogna dichiarare LimitaAvvio nel form_load principale

Public Sub LimitaAvvio(Tipo As Integer, mio As Object, messaggio As String)
On Local Error GoTo errore
If Tipo = 0 Then Exit Sub
If Tipo = 1 Then
    If App.PrevInstance Then
        MsgBox App.Title & " elimino la precedente e riparto"
        Kill App.PrevInstance
        
    End If
    Exit Sub
End If
If Tipo = 2 Then
    If App.PrevInstance Then
    Dim stitle As String
    
    stitle = mio.Caption
    mio.Caption = Hex$(mio.hWnd)
    AppActivate stitle
    End
    End If
    Exit Sub
End If
errore:
End
End Sub


Public Function MakePath(ByVal Path As String) As Boolean

Dim i As Integer, ercode As Long, rc As Boolean
    
    On Error Resume Next
    
    rc = True
    Do
        ' get the next path chunk
        i = InStr(i + 1, Path & "\", "\")
        
        ' try to create this sub-directory
        Err.Clear
        MkDir Left$(Path, i - 1)
        If Err = 0 Then
            ' the directory has been created
            ' do nothing
            
        ElseIf Err = 75 Then
            
            ' Path\File Access Error: the directory exists
            ' do nothing
        Else
            rc = False
            ' we can't continue if any other error
            ercode = Err
            On Error GoTo 0
            Err.Raise ercode
        End If
    Loop Until i > Len(Path)
    MakePath = rc
End Function




Public Function OpenWithDefault(ByVal FileName As String) As Boolean
 'ShellExecute returns a value greater than 32 if it was successful
    OpenWithDefault = (ShellExecute(0&, "", FileName, vbNullString, vbNullString, vbNormalFocus) > 32)
   ' MsgBox OpenWithDefault
End Function







Public Function EnumDir(ByVal PercorsoDirectory As String) As Integer


Dim fso As FileSystemObject
Dim fsoFolder As Folder
Dim fsoFiles As Files
Dim fsoFile As File

Set fso = New FileSystemObject
Set fsoFolder = fso.GetFolder(PercorsoDirectory)
Set fsoFiles = fsoFolder.Files

EnumDir = fsoFiles.Count

End Function

