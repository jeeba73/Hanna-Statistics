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
    
    
Public Function CopyFolder(txtSource As String, txtDestination As String) As Boolean
Dim lFileOp  As Long
Dim lresult  As Long
Dim lFlags   As Long
Dim SHFileOp As SHFILEOPSTRUCT

Screen.MousePointer = vbHourglass

        lFileOp = FO_COPY

'
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
'If lresult <> 0 Or SHFileOp.fAborted Then MsgBox "Problemi di installazione" & vbCrLf & txtDestination, vbInformation, "File Operations"
'
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

Public Function OpenFile(ByVal WhereIsFile As String) As Double

End Function


'Controlla se c'è già una istanza attiva
'se tipo = 0 default avvia una successiva istanza
'se tipo = 1 termina la nuova istanza con un messaggio
'se tipo = 2 passa il controllo alla prima
'N.B. bisogna dichiarare LimitaAvvio nel form_load principale

Public Sub LimitaAvvio(tipo As Integer, mio As Object, messaggio As String)
On Local Error GoTo errore
If tipo = 0 Then Exit Sub
If tipo = 1 Then
    If App.PrevInstance Then
        MsgBox App.Title & " elimino la precedente e riparto"
        Kill App.PrevInstance
        
    End If
    Exit Sub
End If
If tipo = 2 Then
    If App.PrevInstance Then
    Dim sTitle As String
    
    sTitle = mio.Caption
    mio.Caption = Hex$(mio.hwnd)
    AppActivate sTitle
    End
    End If
    Exit Sub
End If
errore:
End
End Sub
Public Function ApriIlReportFolder(ByVal ReportFolder As String) As Boolean
    
Dim rc As Boolean
Dim a As Double
    
    On Error GoTo ERR_APRI
    
    rc = True
    a = Shell("explorer ," & ReportFolder, vbNormalFocus)
    'rc = IIf(a = 2724, True, False)
ERR_END:
    On Error GoTo 0
    ApriIlReportFolder = rc
    Exit Function
ERR_APRI:
  
    Resume ERR_END
End Function

Public Function MakePath(ByVal PATH As String) As Boolean

Dim i As Integer, ercode As Long, rc As Boolean
    
    On Error Resume Next
    
    rc = True
    Do
        ' get the next path chunk
        i = InStr(i + 1, PATH & "\", "\")
        
        ' try to create this sub-directory
        Err.Clear
        MkDir Left$(PATH, i - 1)
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
    Loop Until i > Len(PATH)
    MakePath = rc
End Function

Public Function FormatNomeFile(ByVal MyName As String) As String
    Dim rc As Boolean
    Dim i As Integer
    Dim sString As String
    Dim LeftString As String
    Dim RightString As String
    Dim AccPosition As Integer
    Dim sTrDaModificare(9) As Variant
    
    sTrDaModificare(0) = ":"
    sTrDaModificare(1) = "'"
    sTrDaModificare(2) = ","
    sTrDaModificare(3) = ";"
    sTrDaModificare(4) = "\"
    sTrDaModificare(5) = "/"
    sTrDaModificare(6) = "&"
    sTrDaModificare(7) = "*"
    sTrDaModificare(8) = """"
    sTrDaModificare(9) = " "
        For i = 0 To UBound(sTrDaModificare)
        
            MyName = Replace(MyName, sTrDaModificare(i), ".")

        Next
      
    FormatNomeFile = MyName
    'MsgBox FormatNomeFile
End Function

Public Function FormatID(ByVal MyName As String) As String
    Dim rc As Boolean
    Dim i As Integer
    Dim sString As String
    Dim LeftString As String
    Dim RightString As String
    Dim AccPosition As Integer
    Dim sTrDaModificare(9) As Variant
    
    sTrDaModificare(0) = ":"
    sTrDaModificare(1) = "'"
    sTrDaModificare(2) = ","
    sTrDaModificare(3) = ";"
    sTrDaModificare(4) = "\"
    sTrDaModificare(5) = "/"
    sTrDaModificare(6) = "&"
    sTrDaModificare(7) = "*"
    sTrDaModificare(8) = """"
    sTrDaModificare(9) = " "
    sTrDaModificare(9) = "."
        For i = 0 To UBound(sTrDaModificare)
        
            MyName = Replace(MyName, sTrDaModificare(i), "")

        Next
      
    FormatID = MyName
    'MsgBox FormatNomeFile
End Function





Public Function EnumDir(ByVal PercorsoDirectory As String) As Integer


Dim FSO As FileSystemObject
Dim fsoFolder As Folder
Dim fsoFiles As Files
Dim fsoFile As file

Set FSO = New FileSystemObject
Set fsoFolder = FSO.GetFolder(PercorsoDirectory)
Set fsoFiles = fsoFolder.Files

EnumDir = fsoFiles.Count

End Function

Public Function VerifyFile(FileName As String)
On Error Resume Next
Open FileName For Input As #1
If Err Then
VerifyFile = False
Exit Function
End If
Close #1
VerifyFile = True
End Function




Public Function CheckExsistsFileLots(ByVal bChiuso As Boolean) As Boolean

        Dim rc As Boolean
        Dim PATH As String
        rc = False
        Dim FSO As New Scripting.FileSystemObject
        
        Dim Cartella As Folder
        Dim FileGenerico As file
         
        PATH = IIf(bChiuso, USER_DATA_PATH, USER_TEMP_PATH)
        USER_PATH = PATH
        Set Cartella = FSO.GetFolder(PATH)
         
        For Each FileGenerico In Cartella.Files
        
            If InStr(FileGenerico.Name, USER_ESTENSIONE) Then
                rc = True
           
            End If
        Next

        CheckExsistsFileLots = rc

    
End Function

Public Function SplitPathFile(ByVal szFilename As String, Optional ByRef PATH As String, Optional ByRef Name As String)
    Dim i As Integer
    Dim fields() As String
    If szFilename = "" Then Exit Function
    ' Split the string at the comma characters and add each field to a ListBox
    fields() = Split(szFilename, "\")
    PATH = fields(0)
    For i = 1 To UBound(fields) - 1
        PATH = PATH & "\" & (fields(i))
    Next
    PATH = PATH & "\"
    Name = fields(UBound(fields))

End Function
Public Function SplitName(ByVal szFilename As String, ByRef Name As String)
    Dim i As Integer
    Dim PATH As String
    Dim fields() As String
    
    ' Split the string at the comma characters and add each field to a ListBox
    fields() = Split(szFilename, ":")
    PATH = fields(0)
    For i = 1 To UBound(fields) - 1
        PATH = PATH & ":" & (fields(i))
    Next
    PATH = PATH & ":"
    Name = Trim(fields(UBound(fields)))

End Function
Public Function CheckSavePath(MyPath) As String
    Dim rtn&, pidl&, pos%
    Dim SpecOut As String
    
    If MyPath = "" Then MyPath = App.PATH
    If Right$(MyPath, 1) = "\" Then 'makes sure that "\" is at the end of the path
       SpecOut = MyPath             'if so then, do nothing
    Else                            'otherwise
       SpecOut = MyPath + "\"       'add the "\" to the end of the path
    End If
    CheckSavePath = SpecOut '+ ExtractName(PathStandard) 'merges both the destination path and the source filename into one string
End Function


Function FileIsOpen(sFilename As String) As Boolean
    Dim iFileNum As Integer, lErrNum As Long
    
    On Error Resume Next
    iFileNum = FreeFile()
    'Attempt to open the file and lock it.
    Open sFilename For Input Lock Read As #iFileNum
    Close iFileNum
    lErrNum = Err.NUMBER
    On Error GoTo 0
    
    'Check to see which error occurred.
    Select Case lErrNum
    Case 0
        'No error occurred.
        'File is NOT already open by another user.
        FileIsOpen = False

    Case 70
        'Error number for "Permission Denied."
        'File is already opened by another user.
        FileIsOpen = True
    
    Case 53
        'File not found
        FileIsOpen = False
        
    Case Else
        'Another error occurred.
        FileIsOpen = True
        Debug.Print Error(lErrNum)
    End Select
End Function

