Attribute VB_Name = "modCommonDialog"
Option Explicit

'Choose Folder
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_BROWSEFORCOMPUTER = &H1000

Private Const MAX_PATH = 260

'Open Dialog
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Const GWL_HINSTANCE = (-6)
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NOACTIVATE = &H10
Const HCBT_ACTIVATE = 5
Const WH_CBT = 5

Dim hHook As Long

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORS) As Long
Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONTS) As Long
Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLGS) As Long

Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 256

Public Const LF_FACESIZE = 32

'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not
'a standard Win32 type.
Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY

Public Type OPENFILENAME
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
End Type

Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Type OFNOTIFY
        hdr As NMHDR
        lpOFN As OPENFILENAME
        pszFile As String        '  May be NULL
End Type

Type CHOOSECOLORS
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public Type CHOOSEFONTS
    lStructSize As Long
    hwndOwner As Long          '  caller's window handle
    hdc As Long                '  printer DC/IC or NULL
    lpLogFont As Long          '  ptr. to a LOGFONT struct
    iPointSize As Long         '  10 * size in points of selected font
    flags As Long              '  enum. type flags
    rgbColors As Long          '  returned text color
    lCustData As Long          '  data passed to hook fn.
    lpfnHook As Long           '  ptr. to hook function
    lpTemplateName As String     '  custom template name
    hInstance As Long          '  instance handle of.EXE that
    lpszStyle As String          '  return the style field here
    nFontType As Integer          '  same value reported to the EnumFonts
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long           '  minimum pt size allowed &
    nSizeMax As Long           '  max pt size allowed if
End Type

Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_SHOWHELP = &H8
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_ANYCOLOR = &H100

Public Const COLOR_FLAGS = CC_FULLOPEN Or CC_ANYCOLOR Or CC_RGBINIT

Public Const CF_SCREENFONTS = &H1
Public Const CF_PRINTERFONTS = &H2
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Const CF_SHOWHELP = &H4&
Public Const CF_ENABLEHOOK = &H8&
Public Const CF_ENABLETEMPLATE = &H10&
Public Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_USESTYLE = &H80&
Public Const CF_EFFECTS = &H100&
Public Const CF_APPLY = &H200&
Public Const CF_ANSIONLY = &H400&
Public Const CF_SCRIPTSONLY = CF_ANSIONLY
Public Const CF_NOVECTORFONTS = &H800&
Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Public Const CF_NOSIMULATIONS = &H1000&
Public Const CF_LIMITSIZE = &H2000&
Public Const CF_FIXEDPITCHONLY = &H4000&
Public Const CF_WYSIWYG = &H8000 '  must also have CF_SCREENFONTS CF_PRINTERFONTS
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_SCALABLEONLY = &H20000
Public Const CF_TTONLY = &H40000
Public Const CF_NOFACESEL = &H80000
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_NOSIZESEL = &H200000
Public Const CF_SELECTSCRIPT = &H400000
Public Const CF_NOSCRIPTSEL = &H800000
Public Const CF_NOVERTFONTS = &H1000000

Public Const SIMULATED_FONTTYPE = &H8000
Public Const PRINTER_FONTTYPE = &H4000
Public Const SCREEN_FONTTYPE = &H2000
Public Const BOLD_FONTTYPE = &H100
Public Const ITALIC_FONTTYPE = &H200
Public Const REGULAR_FONTTYPE = &H400

Public Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Public Const SHAREVISTRING = "commdlg_ShareViolation"
Public Const FILEOKSTRING = "commdlg_FileNameOK"
Public Const COLOROKSTRING = "commdlg_ColorOK"
Public Const SETRGBSTRING = "commdlg_SetRGBColor"
Public Const HELPMSGSTRING = "commdlg_help"
Public Const FINDMSGSTRING = "commdlg_FindReplace"

Public Const CD_LBSELNOITEMS = -1
Public Const CD_LBSELCHANGE = 0
Public Const CD_LBSELSUB = 1
Public Const CD_LBSELADD = 2

Type PRINTDLGS
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hdc As Long
        flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long
        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type

Public Const PD_ALLPAGES = &H0
Public Const PD_SELECTION = &H1
Public Const PD_PAGENUMS = &H2
Public Const PD_NOSELECTION = &H4
Public Const PD_NOPAGENUMS = &H8
Public Const PD_COLLATE = &H10
Public Const PD_PRINTTOFILE = &H20
Public Const PD_PRINTSETUP = &H40
Public Const PD_NOWARNING = &H80
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNIC = &H200
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_SHOWHELP = &H800
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLESETUPTEMPLATE = &H8000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NONETWORKBUTTON = &H200000

Type DEVNAMES
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
End Type

Public Const DN_DEFAULTPRN = &H1

Public Type SelectedFile
    nFilesSelected As Integer
    sFiles() As String
    sLastDirectory As String
    bCanceled As Boolean
    
End Type

Public Type SelectedColor
    oSelectedColor As OLE_COLOR
    bCanceled As Boolean
End Type

Public Type SelectedFont
    sSelectedFont As String
    bCanceled As Boolean
    bBold As Boolean
    bItalic As Boolean
    nSize As Integer
    bUnderline As Boolean
    bStrikeOut As Boolean
    lColor As Long
    sFaceName As String
End Type

Public FileDialog As OPENFILENAME
Public ColorDialog As CHOOSECOLORS
Public FontDialog As CHOOSEFONTS
Public PrintDialog As PRINTDLGS
Dim parenthWnd As Long

Private m_CurrentDirectory As String   'The current directory
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long


Public Function ShowOpen(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedFile
Dim ret As Long
Dim Count As Integer
Dim fileNameHolder As String
Dim LastCharacter As Integer
Dim NewCharacter As Integer
Dim tempFiles(1 To 200) As String
Dim hInst As Long
Dim Thread As Long
Dim sSplitArray() As String
Dim sDefaultExt As String
Dim nCount As Integer
    
    parenthWnd = hwnd
    FileDialog.nStructSize = Len(FileDialog)
    FileDialog.hwndOwner = hwnd
    FileDialog.sFileTitle = Space$(2048)
    FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
    If InStr(1, FileDialog.sFile, ".") <> 0 Then
        sDefaultExt = Trim(Mid(FileDialog.sFile, InStr(1, FileDialog.sFile, ".")))
        If InStr(1, UCase(FileDialog.sFilter), UCase(sDefaultExt)) <> 0 Then
            sSplitArray = Split(FileDialog.sFilter, Chr$(0))
            For nCount = 0 To (UBound(sSplitArray)) / 2
'                If InStr(1, UCase(sSplitArray(nCount * 2)), UCase(sDefaultExt)) <> 0 Then
'                    FileDialog.nFilterIndex = nCount + 1
'                    FileDialog.sFile = Mid(FileDialog.sFile, 1, InStr(1, FileDialog.sFile, ".") - 1)
'                    Exit For
'                End If
            Next nCount
        End If
    Else
        If FileDialog.sFilter <> "" Then
            FileDialog.nFilterIndex = 1
        End If
    End If
    FileDialog.sFile = FileDialog.sFile & Space$(2047) & Chr$(0)
    FileDialog.nFileSize = Len(FileDialog.sFile)
    
    If FileDialog.flags = 0 Then
        FileDialog.flags = OFS_FILE_OPEN_FLAGS
    End If
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = GetOpenFileName(FileDialog)

    If ret Then
        If Trim$(FileDialog.sFileTitle) = "" Then
            LastCharacter = 0
            Count = 0
            While ShowOpen.nFilesSelected = 0
                NewCharacter = InStr(LastCharacter + 1, FileDialog.sFile, Chr$(0), vbTextCompare)
                If Count > 0 Then
                    tempFiles(Count) = Mid(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                Else
                    ShowOpen.sLastDirectory = Mid(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                End If
                Count = Count + 1
                If InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0), vbTextCompare) = InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0) & Chr$(0), vbTextCompare) Then
                    tempFiles(Count) = Mid(FileDialog.sFile, NewCharacter + 1, InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0) & Chr$(0), vbTextCompare) - NewCharacter - 1)
                    ShowOpen.nFilesSelected = Count
                End If
                LastCharacter = NewCharacter
            Wend
            ReDim ShowOpen.sFiles(1 To ShowOpen.nFilesSelected)
            For Count = 1 To ShowOpen.nFilesSelected
                ShowOpen.sFiles(Count) = tempFiles(Count)
            Next
        Else
            ReDim ShowOpen.sFiles(1 To 1)
            ShowOpen.sLastDirectory = Left$(FileDialog.sFile, FileDialog.nFileOffset)
            ShowOpen.nFilesSelected = 1
            ShowOpen.sFiles(1) = Mid(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFile, Chr$(0), vbTextCompare) - FileDialog.nFileOffset - 1)
        End If
        ShowOpen.bCanceled = False
        Exit Function
    Else
        ShowOpen.sLastDirectory = ""
        ShowOpen.nFilesSelected = 0
        ShowOpen.bCanceled = True
        Erase ShowOpen.sFiles
        Exit Function
    End If
End Function

Public Function ShowSave(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedFile
Dim ret As Long
Dim hInst As Long
Dim Thread As Long
Dim sDefaultExtention As String
Dim sSplitArray() As String
Dim sDefaultStartExt As String
Dim nCount As Integer
    
    parenthWnd = hwnd
    FileDialog.nStructSize = Len(FileDialog)
    FileDialog.hwndOwner = hwnd
    FileDialog.sFileTitle = Space$(2048)
    FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
    If InStr(1, FileDialog.sFile, ".") <> 0 Then
        sDefaultStartExt = Trim(Mid(FileDialog.sFile, InStr(1, FileDialog.sFile, ".")))
        If InStr(1, FileDialog.sFilter, sDefaultStartExt) <> 0 Then
            sSplitArray = Split(FileDialog.sFilter, Chr$(0))
            For nCount = 0 To (UBound(sSplitArray) + 1) / 2
                If InStr(1, sSplitArray(nCount * 2), sDefaultStartExt) <> 0 Then
                    FileDialog.nFilterIndex = nCount + 1
                    FileDialog.sFile = Mid(FileDialog.sFile, 1, InStr(1, FileDialog.sFile, ".") - 1)
                    Exit For
                End If
            Next nCount
        End If
    Else
        If FileDialog.sFilter <> "" Then
            FileDialog.nFilterIndex = 1
        End If
    End If
    FileDialog.sFile = Space$(2047) & Chr$(0)
    FileDialog.nFileSize = Len(FileDialog.sFile)
    
    If FileDialog.flags = 0 Then
        FileDialog.flags = OFS_FILE_SAVE_FLAGS
    End If
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = GetSaveFileName(FileDialog)
    ReDim ShowSave.sFiles(1)

    If ret Then
        ShowSave.sLastDirectory = Left$(FileDialog.sFile, FileDialog.nFileOffset)
        ShowSave.nFilesSelected = 1
        ShowSave.sFiles(1) = Mid(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFile, Chr$(0), vbTextCompare) - FileDialog.nFileOffset - 1)
        If InStr(1, FileDialog.sFilter, Chr$(0)) <> 0 Then
            sSplitArray = Split(FileDialog.sFilter, Chr$(0))
            sDefaultExtention = Mid(sSplitArray(FileDialog.nFilterIndex * 2 - 1), InStr(1, sSplitArray(FileDialog.nFilterIndex * 2 - 1), "."))
            If InStr(1, ShowSave.sFiles(1), ".") = 0 Then
                ShowSave.sFiles(1) = ShowSave.sFiles(1) & sDefaultExtention
            ElseIf InStr(1, ShowSave.sFiles(1), "." & sDefaultExtention) = 0 And UBound(sSplitArray) = 2 Then
                ShowSave.sFiles(1) = Mid(ShowSave.sFiles(1), 1, InStr(1, ShowSave.sFiles(1), ".") - 1) & sDefaultExtention
            End If
        End If
        ShowSave.bCanceled = False
        Exit Function
    Else
        ShowSave.sLastDirectory = ""
        ShowSave.nFilesSelected = 0
        ShowSave.bCanceled = True
        Erase ShowSave.sFiles
        Exit Function
    End If
End Function

Public Function ShowColor(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedColor
Dim CustomColors() As Byte  ' dynamic (resizable) array
Dim i As Integer
Dim ret As Long
Dim hInst As Long
Dim Thread As Long

    parenthWnd = hwnd
    If ColorDialog.lpCustColors = "" Then
        ReDim CustomColors(0 To 16 * 4 - 1) As Byte  'resize the array
    
        For i = LBound(CustomColors) To UBound(CustomColors)
          CustomColors(i) = 254 ' sets all custom colors to white
        Next i
        
        ColorDialog.lpCustColors = StrConv(CustomColors, vbUnicode)  ' convert array
    End If
    
    ColorDialog.hwndOwner = hwnd
    ColorDialog.lStructSize = Len(ColorDialog)
    ColorDialog.flags = COLOR_FLAGS
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = ChooseColor(ColorDialog)
    If ret Then
        ShowColor.bCanceled = False
        ShowColor.oSelectedColor = ColorDialog.rgbResult
        Exit Function
    Else
        ShowColor.bCanceled = True
        ShowColor.oSelectedColor = &H0&
        Exit Function
    End If
End Function

Public Function ShowFont(ByVal hwnd As Long, ByVal startingFont As StdFont, Optional ByVal centerForm As Boolean = True, Optional fontColor As OLE_COLOR) As SelectedFont
Dim ret As Long
Dim lfLogFont As LOGFONT
Dim hInst As Long
Dim Thread As Long
Dim i As Integer
    
    parenthWnd = hwnd
    With startingFont
        lfLogFont.lfStrikeOut = .Strikethrough
        lfLogFont.lfUnderline = .Underline
        lfLogFont.lfWeight = IIf(.Bold = True, 700, 0)
        lfLogFont.lfItalic = .Italic
        lfLogFont.lfHeight = CInt(startingFont.Size) * 1.347
    End With
    FontDialog.nSizeMax = 0
    FontDialog.nSizeMin = 0
    FontDialog.nFontType = Screen.FontCount
    FontDialog.hwndOwner = hwnd
    FontDialog.hdc = 0
    FontDialog.lpfnHook = 0
    FontDialog.lCustData = 0
    FontDialog.lpLogFont = VarPtr(lfLogFont)
    FontDialog.lpTemplateName = Space$(2048)
    FontDialog.rgbColors = fontColor
    FontDialog.iPointSize = 10
    FontDialog.lStructSize = Len(FontDialog)
    
    If FontDialog.flags = 0 Then
        FontDialog.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT 'Or CF_EFFECTS
    End If
    
    For i = 0 To Len(startingFont.Name) - 1
        lfLogFont.lfFaceName(i) = Asc(Mid(startingFont.Name, i + 1, 1))
    Next
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = ChooseFont(FontDialog)
        
    If ret Then
        ShowFont.bCanceled = False
        ShowFont.bBold = IIf(lfLogFont.lfWeight > 400, 1, 0)
        ShowFont.bItalic = lfLogFont.lfItalic
        ShowFont.bStrikeOut = lfLogFont.lfStrikeOut
        ShowFont.bUnderline = lfLogFont.lfUnderline
        ShowFont.lColor = FontDialog.rgbColors
        ShowFont.nSize = FontDialog.iPointSize / 10
        For i = 0 To 31
            ShowFont.sSelectedFont = ShowFont.sSelectedFont + Chr(lfLogFont.lfFaceName(i))
        Next
    
        ShowFont.sSelectedFont = Mid(ShowFont.sSelectedFont, 1, InStr(1, ShowFont.sSelectedFont, Chr(0)) - 1)
        Exit Function
    Else
        ShowFont.bCanceled = True
        Exit Function
    End If
End Function
Public Function ShowPrinter(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As Long
Dim hInst As Long
Dim Thread As Long
    
    parenthWnd = hwnd
    PrintDialog.hwndOwner = hwnd
    PrintDialog.lStructSize = Len(PrintDialog)
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ShowPrinter = PrintDlg(PrintDialog)
End Function

Public Function ShowFolder(ByVal hwnd As Long, Optional Title As String, Optional startingDir As String, Optional ByVal centerForm As Boolean = True) As String
'Opens a Treeview control that displays
'     the directories in a computer
Dim hInst As Long
Dim Thread As Long
Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

    m_CurrentDirectory = startingDir & vbNullChar

    parenthWnd = hwnd
    If Title <> "" Then
        szTitle = Title
    Else
        szTitle = "Please Select a Folder"
    End If

    With tBrowseInfo
        .hwndOwner = hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
        .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)  'get address of function.
    End With
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)


    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        ShowFolder = sBuffer
    End If
End Function
Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
  
  Dim lpIDList As Long
  Dim ret As Long
  Dim sBuffer As String
  
  On Error Resume Next  'Sugested by MS to prevent an error from
                        'propagating back into the calling process.
     
  Select Case uMsg
  
    Case BFFM_INITIALIZED
      Call SendMessage(hwnd, BFFM_SETSELECTION, 1, m_CurrentDirectory)
      
    Case BFFM_SELCHANGED
      sBuffer = Space(MAX_PATH)
      
      ret = SHGetPathFromIDList(lp, sBuffer)
      If ret = 1 Then
        Call SendMessage(hwnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
      End If
      
  End Select
  
  BrowseCallbackProc = 0
  
End Function
' This function allows you to assign a function pointer to a vaiable.
Private Function GetAddressofFunction(add As Long) As Long
  GetAddressofFunction = add
End Function

Private Function WinProcCenterScreen(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim rectForm As RECT, rectMsg As RECT
    Dim X As Long, Y As Long
    If lMsg = HCBT_ACTIVATE Then
        'Show the MsgBox at a fixed location (0,0)
        GetWindowRect wParam, rectMsg
        X = Screen.Width / Screen.TwipsPerPixelX / 2 - (rectMsg.Right - rectMsg.Left) / 2
        Y = Screen.Height / Screen.TwipsPerPixelY / 2 - (rectMsg.Bottom - rectMsg.Top) / 2
        SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
    End If
    WinProcCenterScreen = False
End Function

Private Function WinProcCenterForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim rectForm As RECT, rectMsg As RECT
    Dim X As Long, Y As Long
    'On HCBT_ACTIVATE, show the MsgBox centered over Form1
    If lMsg = HCBT_ACTIVATE Then
        'Get the coordinates of the form and the message box so that
        'you can determine where the center of the form is located
        GetWindowRect parenthWnd, rectForm
        GetWindowRect wParam, rectMsg
        X = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - ((rectMsg.Right - rectMsg.Left) / 2)
        Y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - ((rectMsg.Bottom - rectMsg.Top) / 2)
        'Position the msgbox
        SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
     End If
     WinProcCenterForm = False
End Function

Public Function DetermineDirectory(inputString As String, Optional includeLastSlash As Boolean = True) As String
Dim pos As Integer
    pos = InStrRev(inputString, "\", , vbTextCompare)
    DetermineDirectory = Mid(inputString, 1, pos)
    If includeLastSlash = False Then
        DetermineDirectory = Mid(DetermineDirectory, 1, Len(DetermineDirectory) - 1)
    End If
End Function
Public Function DetermineFilename(inputString As String) As String
Dim pos As Integer
    pos = InStrRev(inputString, "\", , vbTextCompare)
    DetermineFilename = Mid(inputString, pos + 1, Len(inputString) - pos)
End Function

Public Function SelectFile(parenthWnd As Long, Title As String, filter As String, flags As Long, defaultExtension As String, Optional saveFile As Boolean = True, Optional lastFilename As String) As String
Dim sOpen As SelectedFile
Dim filename As String
Dim ret As Integer

    On Error GoTo e_Browse
    FileDialog.sFilter = filter
    FileDialog.flags = flags
    FileDialog.sDlgTitle = Title
    If lastFilename <> "" Then
        FileDialog.sInitDir = DetermineDirectory(lastFilename)
        FileDialog.sFile = DetermineFilename(lastFilename)
    ElseIf FileDialog.sInitDir = "" Then
        FileDialog.sInitDir = App.Path
        FileDialog.sFile = ""
    Else
        FileDialog.sFile = ""
    End If
    
    Do While filename = ""
        If saveFile = False Then
            sOpen = ShowOpen(parenthWnd, True)
        Else
            sOpen = ShowSave(parenthWnd, True)
        End If
        If sOpen.sFiles(1) = "" Then
            ret = MsgBox("Please select a " & Title, vbOKCancel + vbInformation, "Missing Filename")
            If ret = vbCancel Then
                Exit Function
            End If
        Else
            filename = sOpen.sLastDirectory & sOpen.sFiles(1)
            If InStr(1, filename, ".") = 0 Then
                If LCase(Right(filename, 4)) <> "." & defaultExtension Then
                    filename = filename & "." & defaultExtension
                End If
            End If
            SelectFile = filename
        End If
    Loop
    Exit Function
e_Browse:
    SelectFile = ""
    Exit Function
End Function
