Attribute VB_Name = "CommDlgs"
'//
'// Common Dialogs Modulo
'//
'// Descrizione:
'// Richiama le finestre di dialogo del sistema,
'// senza l'uso di OCX o DLL di terze parti.
'//

Option Explicit

' Strutture
Public szFileTitle
Private Type OPENFILENAME
    lStructSize As Long
    hWnd As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type COLORSTRUC
    lStructSize As Long
    hWnd As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const LF_FACESIZE = 32

Private Type LOGFONT
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

Private Type FONTSTRUC
    lStructSize As Long
    hWnd As Long
    hdc As Long
    lpLogFont As Long
    iPointSize As Long
    flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Public Type DevMode
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFreq As Long
End Type

Private Type PRINTDLGSTRUC
    lStructSize As Long
    hWnd As Long
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

Public Type PRINTPROPS
    Cancel As Boolean
    Device As String
    Copies As Integer
    Collate As Boolean
    file As Boolean
    All As Boolean
    Pages As Boolean
    Selection As Boolean
    FromPage As Integer
    ToPage As Integer
    dm As DevMode
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Type BrowseInfo
    hOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

' Dichiarazione funzioni (private)

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLGSTRUC) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As COLORSTRUC) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As FONTSTRUC) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long 'ITEMIDLIST
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

' Dichiarazione funzioni (pubbliche)

Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hWnd As Long, ByVal szFilename As String, ByVal dwCommand As Long, ByVal dwData As Any) As Long

' Costanti  Stampa Dialog Box

Public Const PD_NOSELECTION = &H4
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_PRINTTOFILE = &H20
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_RETURNIC = &H200
Public Const PD_SELECTION = &H1
Public Const PD_SHOWHELP = &H800
Public Const PD_NOPAGENUMS = &H8
Public Const PD_PAGENUMS = &H2
Public Const PD_ALLPAGES = &H0
Public Const PD_COLLATE = &H10
Public Const PD_HIDEPRINTTOFILE = &H100000

' Costanti  WinHelp

Public Const HELP_COMMAND = &H102&
Public Const HELP_CONTENTS = &H3&
Public Const HELP_CONTEXT = &H1
Public Const HELP_CONTEXTPOPUP = &H8&
Public Const HELP_FORCEFILE = &H9&
Public Const HELP_HELPONHELP = &H4
Public Const HELP_INDEX = &H3
Public Const HELP_KEY = &H101
Public Const HELP_MULTIKEY = &H201&
Public Const HELP_PARTIALKEY = &H105&
Public Const HELP_QUIT = &H2
Public Const HELP_SETCONTENTS = &H5&
Public Const HELP_SETINDEX = &H5
Public Const HELP_SETWINPOS = &H203&

' Costanti  HTMLHelp

Public Const HH_DISPLAY_TOPIC = &H0&
Public Const HH_HELP_FINDER = &H0&
Public Const HH_DISPLAY_TOC = &H1&    'Non Implementata
Public Const HH_DISPLAY_INDEX = &H2&  'Non Implementata
Public Const HH_DISPLAY_SEARCH = &H3& 'Non Implementata
Public Const HH_SET_WIN_TYPE = &H4&
Public Const HH_GET_WIN_TYPE = &H5&
Public Const HH_GET_WIN_HANDLE = &H6&
Public Const HH_ENUM_INFO_TYPE = &H7&
Public Const HH_SET_INFO_TYPE = &H8&
Public Const HH_SYNC = &H9&
Public Const HH_ADD_NAV_UI = &H10&     '// Currently Not Implemented
Public Const HH_ADD_BUTTON = &H11&     '// Currently Not Implemented
Public Const HH_GETBROWSER_APP = &H12& '// Currently Not Implemented
Public Const HH_KEYWORD_LOOKUP = &H13&
Public Const HH_DISPLAY_TEXT_POPUP = &H14&
Public Const HH_HELP_CONTEXT = &H15&
Public Const HH_TP_HELP_CONTEXTMENU = &H16&
Public Const HH_TP_HELP_WM_HELP = &H17&
Public Const HH_CLOSE_ALL = &H18&
Public Const HH_ALINK_LOOKUP = &H19&
Public Const HH_GET_LAST_ERROR = &H20&       '// Currently Not Implemented
Public Const HH_ENUM_CATEGORY = &H21&
Public Const HH_ENUM_CATEGORY_IT = &H22&
Public Const HH_RESET_IT_FILTER = &H23&
Public Const HH_SET_INCLUSIVE_FILTER = &H24&
Public Const HH_SET_EXCLUSIVE_FILTER = &H25&
Public Const HH_SET_GUID = &H26&
Public Const HH_INTERNAL = &H255&

' Costanti Private

Private Const FW_BOLD = 700
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10
Private Const PD_ENABLEPRINTHOOK = &H1000
Private Const PD_ENABLEPRINTTEMPLATE = &H4000
Private Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Private Const PD_ENABLESETUPHOOK = &H2000
Private Const PD_ENABLESETUPTEMPLATE = &H8000
Private Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Private Const PD_NONETWORKBUTTON = &H200000
Private Const PD_PRINTSETUP = &H40
Private Const PD_USEDEVMODECOPIES = &H40000
Private Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Private Const PD_NOWARNING = &H80
Private Const CF_ANSIONLY = &H400&
Private Const CF_APPLY = &H200&
Private Const CF_BITMAP = 2
Private Const CF_PRINTERFONTS = &H2
Private Const CF_PRIVATEFIRST = &H200
Private Const CF_PRIVATELAST = &H2FF
Private Const CF_RIFF = 11
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_SCREENFONTS = &H1
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_DIB = 8
Private Const CF_DIF = 5
Private Const CF_DSPBITMAP = &H82
Private Const CF_DSPENHMETAFILE = &H8E
Private Const CF_DSPMETAFILEPICT = &H83
Private Const CF_DSPTEXT = &H81
Private Const CF_EFFECTS = &H100&
Private Const CF_ENABLEHOOK = &H8&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLETEMPLATEHANDLE = &H20&
Private Const CF_ENHMETAFILE = 14
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_GDIOBJFIRST = &H300
Private Const CF_GDIOBJLAST = &H3FF
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_METAFILEPICT = 3
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOVERTFONTS = &H1000000
Private Const CF_NOVECTORFONTS = &H800&
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_NOSIMULATIONS = &H1000&
Private Const CF_NOSIZESEL = &H200000
Private Const CF_NOSTYLESEL = &H100000
Private Const CF_OEMTEXT = 7
Private Const CF_OWNERDISPLAY = &H80
Private Const CF_PALETTE = 9
Private Const CF_PENDATA = 10
Private Const CF_SCRIPTSONLY = CF_ANSIONLY
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_SHOWHELP = &H4&
Private Const CF_SYLK = 4
Private Const CF_TEXT = 1
Private Const CF_TIFF = 6
Private Const CF_TTONLY = &H40000
Private Const CF_UNICODETEXT = 13
Private Const CF_USESTYLE = &H80&
Private Const CF_WAVE = 12
Private Const CF_WYSIWYG = &H8000
Private Const CFERR_CHOOSEFONTCODES = &H2000
Private Const CFERR_MAXLESSTHANMIN = &H2002
Private Const CFERR_NOFONTS = &H2001
Private Const CC_ANYCOLOR = &H100
Private Const CC_CHORD = 4
Private Const CC_CIRCLES = 1
Private Const CC_ELLIPSES = 8
Private Const CC_ENABLEHOOK = &H10
Private Const CC_ENABLETEMPLATE = &H20
Private Const CC_ENABLETEMPLATEHANDLE = &H40
Private Const CC_FULLOPEN = &H2
Private Const CC_INTERIORS = 128
Private Const CC_NONE = 0
Private Const CC_PIE = 2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_RGBINIT = &H1
Private Const CC_ROUNDRECT = 256 '
Private Const CC_SHOWHELP = &H8
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_STYLED = 32
Private Const CC_WIDE = 16
Private Const CC_WIDESTYLED = 64
Private Const CCERR_CHOOSECOLORCODES = &H5000
Private Const LOGPIXELSY = 90
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const SIMULATED_FONTTYPE = &H8000
Private Const PRINTER_FONTTYPE = &H4000
Private Const SCREEN_FONTTYPE = &H2000
Private Const BOLD_FONTTYPE = &H100
Private Const ITALIC_FONTTYPE = &H200
Private Const REGULAR_FONTTYPE = &H400
Private Const WM_CHOOSEFONT_GETLOGFONT = (&H400 + 1)
Private Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Private Const SHAREVISTRING = "commdlg_ShareViolation"
Private Const FILEOKSTRING = "commdlg_FileNameOK"
Private Const COLOROKSTRING = "commdlg_ColorOK"
Private Const SETRGBSTRING = "commdlg_SetRGBColor"
Private Const FINDMSGSTRING = "commdlg_FindReplace"
Private Const HELPMSGSTRING = "commdlg_help"
Private Const CD_LBSELNOITEMS = -1
Private Const CD_LBSELCHANGE = 0
Private Const CD_LBSELSUB = 1
Private Const CD_LBSELADD = 2
Private Const NOERROR = 0
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_TEMPLATES = &H15
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATOTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_WININICHANGE = &H1A

'// SetDefaultPrinter
'//
'// Descrizione:
'// Imposta la stampante predefinita.
'//
'// Sintassi:
'// BOOL = SetDefaultPrinter(object)
'//
'// Esempio:
'// Dim objNewPrinter As Printer
'// Set objNewPrinter = Printers(2)
'// SetDefaultPrinter objNewPrinter
'//

Public Function SetDefaultPrinter(objPrn As Printer) As Boolean

    Dim X As Long, szTmp As String
    
    szTmp = objPrn.DeviceName & "," & objPrn.DriverName & "," & objPrn.Port
    X = WriteProfileString("windows", "device", szTmp)
    X = SendMessageByString(HWND_BROADCAST, WM_WININICHANGE, 0&, "windows")
    
End Function

'//
'// GetDefaultPrinter
'//
'// Descrizione:
'// Ritorna il nome della stampante predefinita
'//
'// Syntassi:
'// StrVar = GetDefaultPrinter()
'//
'// Esempio:
'// szDefPrinter = GetDefaultPrinter
'//

Public Function GetDefaultPrinter() As String

    Dim X As Long, szTmp As String, dwBuf As Long

    dwBuf = 1024
    szTmp = Space(dwBuf + 1)
    X = GetProfileString("windows", "device", "", szTmp, dwBuf)
    GetDefaultPrinter = Trim(Left(szTmp, X))

End Function

'//
'// ResetDefaultPrinter
'//
'// Descrizione:
'// Resetta la stampante predefinita
'//
'// Syntassi:
'// BOOL = ResetDefaultPrinter(StrVar)
'//
'// Esempio:
'// szDefPrinter = GetDefaultPrinter()
'// If Not ResetDefaultPrinter(szDefPrinter) Then
'//     MsgBox "Non posso resettare la stampante predefinita.", vbExclamation
'// End If
'//

Public Function ResetDefaultPrinter(szBuf As String) As Boolean

    Dim X As Long
    
    X = WriteProfileString("windows", "device", szBuf)
    X = SendMessageByString(HWND_BROADCAST, WM_WININICHANGE, 0&, "windows")

End Function

'//
'// BrowseFolder
'//
'// Descrizione:
'// Apre la finestra Sfoglia cartelle
'//
'// Syntassi:
'// StrVar = BrowseFolder(hWnd, StrVar)
'//
'// Esempio:
'// szFilename = BrowseFolder(Me.hWnd, "Seleziona la cartella desiderata:")
'//

Public Function BrowseFolder(hWnd As Long, szDialogTitle As String) As String

    Dim X As Long, BI As BrowseInfo, dwIList As Long, szPath As String, wPos As Integer
    
    BI.hOwner = hWnd
    BI.lpszTitle = szDialogTitle
    BI.ulFlags = BIF_RETURNONLYFSDIRS
    dwIList = SHBrowseForFolder(BI)
    szPath = Space$(512)
    X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    If X Then
        wPos = InStr(szPath, Chr(0))
        BrowseFolder = Left$(szPath, wPos - 1)
    Else
        BrowseFolder = ""
    End If

End Function

'//
'// DialogConnectToPrinter
'//
'// Descrizione:
'// Per selezionare la stampante che interessa.
'//
'// Syntassi:
'// DialogConnectToPrinter
'//
'// Esempio:
'// DialogConnectToPrinter
'//

Public Function DialogConnectToPrinter() As Boolean

    Shell "rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL AddPrinter", vbNormalFocus
    
End Function

'//
'// ByteToString
'//
'// Descrizione:
'// Converte un array di bytes in una stringa
'//
'// Syntassi:
'// StrVar = ByteToString(ARRAY)
'//
'// Esempio:
'// szBuf = BytesToString(aChars(10))
'//

Private Function ByteToString(aBytes() As Byte) As String

    Dim dwBytePoint As Long, dwByteVal As Long, szOut As String
    
    dwBytePoint = LBound(aBytes)
    
    While dwBytePoint <= UBound(aBytes)
        
        dwByteVal = aBytes(dwBytePoint)
        
        If dwByteVal = 0 Then
            ByteToString = szOut
            Exit Function
        Else
            szOut = szOut & Chr$(dwByteVal)
        End If
        
        dwBytePoint = dwBytePoint + 1
    
    Wend
    
    ByteToString = szOut
    
End Function

'//
'// DialogColor
'//
'// Descrizione:
'// Apre la finestra di dialogo per la scelta del colore.
'//
'// Syntassi:
'// BOOL = DialogColor(hWnd, CONTROL)
'//
'// Esempio:
'// Dim yn as Boolean
'// yn = DialogColor(Me.hWnd, txtEditor)
'//

Public Function DialogColor(hWnd As Long, c As Control) As Boolean

    Dim X As Long, CS As COLORSTRUC, CustColor(16) As Long
    
    CS.lStructSize = Len(CS)
    CS.hWnd = hWnd
    CS.hInstance = App.hInstance
    CS.flags = CC_SOLIDCOLOR
    CS.lpCustColors = String$(16 * 4, 0)
    X = ChooseColor(CS)
    If X = 0 Then
        DialogColor = False
    Else
        DialogColor = True
        c.ForeColor = CS.rgbResult
    End If
    
End Function

'//
'// DialogFile
'//
'// Descrizione:
'// Apre la finestra di dialogo Apri/Salva.
'//
'// Syntassi:
'// StrVar = DialogFile(hWnd, IntVar, StrVar, StrVar, StrVar, StrVar, StrVar)
'//
'// Esempio:
'// szFilename = DialogFile(Me.hWnd, 1, "Open", "MyFileName.doc", "Documents" & Chr(0) & "*.doc" & Chr(0) & "Tutti i files" & Chr(0) & "*.*", DirArchivioBack, "doc")
'//
'// Nota che la variabile szFilter lavora differentemente
'// a seconda del filtro associato.
'// Occorre separare i filtri col carattere |

Public Function DialogFile(hWnd As Long, wMode As Integer, szDialogTitle As String, szFilename As String, szFilter As String, szDefDir As String, szDefExt As String) As String

    Dim X As Long, OFN As OPENFILENAME, szFile As String, szFileTitle As String

    OFN.lStructSize = Len(OFN)
    OFN.hWnd = hWnd
    OFN.lpstrTitle = szDialogTitle
    OFN.lpstrFile = szFilename & String$(250 - Len(szFilename), 0)
    OFN.nMaxFile = 255
    OFN.lpstrFileTitle = String$(255, 0)
    OFN.nMaxFileTitle = 255
    OFN.lpstrFilter = szFilter
    OFN.nFilterIndex = 1
    OFN.lpstrInitialDir = szDefDir
    OFN.lpstrDefExt = szDefExt

    If wMode = 1 Then
        OFN.flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
        X = GetOpenFileName(OFN)
    Else
        OFN.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
        X = GetSaveFileName(OFN)
    End If
    If X <> 0 Then
    
        '// If InStr(OFN.lpstrFileTitle, Chr$(0)) > 0 Then
        '//     szFileTitle = Left$(OFN.lpstrFileTitle, InStr(OFN.lpstrFileTitle, Chr$(0)) - 1)
        '// End If
        If InStr(OFN.lpstrFile, Chr$(0)) > 0 Then
            szFile = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, Chr$(0)) - 1)
        End If
       
        '// OFN.nFileOffset è il numero di caratteri accettato
        '// inizialmente dal nome del file compreso il percorso
        '// OFN.nFileExtension è il numero di caratteri dell'estensione
        '// incluso (.)
        '// MsgBox "Il nome del file è " & szFileTitle & Chr$(13) & Chr$(10) & "Percorso e file é " & szFile, , "Open"
        '// DialogFile = szFile & "|" & szFileTitle
        DialogFile = szFile
    Else
    
        DialogFile = ""
        
    End If
    
End Function

'//
'// DialogFont
'//
'// Descrizione:
'// Apre la finestra di dialogo per i caratteri
'//
'// Syntassi:
'// BOOL = DialogFont(hWnd, CONTROL)
'//
'// Esempio:
'// Dim yn as Boolean
'// yn = DialogFont(Me.hWnd, txtEditor)
'//

Public Function DialogFont(hWnd As Long, c As Control) As String

    Dim lf As LOGFONT, FS As FONTSTRUC
    Dim lLogFontAddress As Long, lMemHandle As Long
    
    If c.Font.Bold Then lf.lfWeight = FW_BOLD
    If c.Font.Italic = True Then lf.lfItalic = 1
    If c.Font.Underline = True Then lf.lfUnderline = 1
    
    FS.lStructSize = Len(FS)
    
    lMemHandle = GlobalAlloc(GHND, Len(lf))
    If lMemHandle = 0 Then
        DialogFont = False
        Exit Function
    End If
    
    lLogFontAddress = GlobalLock(lMemHandle)
    If lLogFontAddress = 0 Then
        DialogFont = False
        Exit Function
    End If
    
    CopyMemory ByVal lLogFontAddress, lf, Len(lf)
    FS.lpLogFont = lLogFontAddress
    FS.iPointSize = c.Font.Size * 10
    FS.flags = CF_SCREENFONTS Or CF_EFFECTS
    
    If ChooseFont(FS) = 1 Then
    
        CopyMemory lf, ByVal lLogFontAddress, Len(lf)
            
        If lf.lfWeight >= FW_BOLD Then
            c.Font.Bold = True
        Else
            c.Font.Bold = False
        End If
                        
        If lf.lfItalic = 1 Then
            c.Font.Italic = True
        Else
            c.Font.Italic = False
        End If
            
        If lf.lfUnderline = 1 Then
            c.Font.Underline = True
        Else
            c.Font.Underline = False
        End If
        On Error GoTo ERR_DIALOG:
        c.Font.Name = ByteToString(lf.lfFaceName())
        c.Font.Size = CLng(FS.iPointSize / 10)
        
        DialogFont = c.Font.Name
        'UserSize = c.Font.Size
            
    Else
    
        '0DialogFont = False
            
    End If
ERR_RESUME:
    On Error GoTo 0
    Exit Function
ERR_DIALOG:
    'DialogFont = False
    Resume ERR_RESUME
End Function

'//
'// DialogPrint
'//
'// Descrizione:
'// Apre la finestra di dialogo Stampa
'//
'// Syntassi:
'// PRINTPROPS = DialogPrint(hWnd, BOOL, DWORD)
'//
'// Esempio:
'// Dim PP As PRINTPROPS
'// PP = DialogPrint(Me.hWnd, True, PD_PAGENUMS or PD_SELECTION or PD_SHOWHELP)
'//

Public Function DialogPrint(hWnd As Long, bPages As Boolean, flags As Long) As PRINTPROPS

    Dim dm As DevMode, pd As PRINTDLGSTRUC
    Dim lpDM As Long, wNull As Integer, szDevName As String
    
    pd.lStructSize = Len(pd)
    pd.hWnd = hWnd
    pd.hDevMode = 0
    pd.hDevNames = 0
    pd.hdc = 0
    pd.flags = flags
    pd.nFromPage = 0
    pd.nToPage = 0
    pd.nMinPage = 0
    If bPages Then pd.nMaxPage = bPages - 1
    pd.nCopies = 0
    DialogPrint.Cancel = True
    
    If PrintDlg(pd) Then
    
        lpDM = GlobalLock(pd.hDevMode)
        CopyMemory dm, ByVal lpDM, Len(dm)
        lpDM = GlobalUnlock(pd.hDevMode)
        
        DialogPrint.Cancel = False
        DialogPrint.Device = Left$(dm.dmDeviceName, InStr(dm.dmDeviceName, Chr(0)) - 1)
        DialogPrint.FromPage = 0
        DialogPrint.ToPage = 0
        DialogPrint.All = True
        If pd.flags And PD_PRINTTOFILE Then DialogPrint.file = True Else DialogPrint.file = False
        If pd.flags And PD_COLLATE Then DialogPrint.Collate = True Else DialogPrint.Collate = False
        If pd.flags And PD_PAGENUMS Then
            DialogPrint.Pages = True
            DialogPrint.All = False
            DialogPrint.FromPage = pd.nFromPage
            DialogPrint.ToPage = pd.nToPage
        Else
            DialogPrint.Pages = False
        End If
        If pd.flags And PD_SELECTION Then
            DialogPrint.Selection = True
            DialogPrint.All = False
        Else
            DialogPrint.Pages = False
        End If
        
        If pd.nCopies = 1 Then
            DialogPrint.Copies = dm.dmCopies
        End If
        
        DialogPrint.dm = dm
        
    End If
    
End Function

'//
'// DialogPrintSetup
'//
'// Descrizione:
'// Apre la finestra di dialogo Imposta stampante
'//
'// Syntassi:
'// BOOL = DialogPrintSetup(hWnd)
'//
'// Esempio:
'// If DialogPrintSetup(Me.hWnd) Then
'// End If
'//

Public Function DialogPrintSetup(hWnd As Long) As Boolean
    Dim X As Long, pd As PRINTDLGSTRUC
    pd.lStructSize = Len(pd)
    pd.hWnd = hWnd
    pd.flags = PD_PRINTSETUP
    X = PrintDlg(pd)
End Function
