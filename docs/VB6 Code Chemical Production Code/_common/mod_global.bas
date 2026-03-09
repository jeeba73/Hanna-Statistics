Attribute VB_Name = "mod_global"
Option Explicit



Public bOpenProductClassificationAfterScan As Boolean


Public MainWindowState As Integer
'------------------------------------------
'       Variabili Globali
'------------------------------------------

' per spostare la FORM

Public FrmMove As Boolean
Public DragX, DragY   As Single

'------------------------------------------

Public bFullScreen              As Boolean

Public SettingName              As String
Public UsedFileName             As String
Public PdfFileName              As String
Public bSearchClosedLot        As Boolean



'------------------------------------------
'       Costanti
'------------------------------------------

Public Const USER_ESTENSIONE = ".cp"
Public Const USER_ESTENSIONE_RFP = ".rfp"
Public Const USER_ESTENSIONE_PREPARATION = ".prep"
Public Const USER_ESTENSIONE_PRODUCTION = ".prod"
Public Const ProjectName = "ChemicalProduction"

Public Const MaxTipoLetture = 10

'------------------------------------------
'       colori standard
'------------------------------------------

          
Public Const vbColorRosaTabella = &HC7B8FE
Public Const vbColorRed = &HC0&
Public Const vbColorAzzurrino = &HFFE1C1
Public Const vbColorResults = &HD0D0C0
Public Const vbColorIns = &HE0E0E0
Public Const vbColorgotFocus = &HF0F0F0
Public Const vbColorUnabled = &HC0C0C0
Public Const vbColorDarkUnabled = &H202020
Public Const vbColorOrange = &H80FF&
Public Const vbColorDarkFont = &H404040
Public Const vbColorGreen = &H8000&    '&H4000&
Public Const vbColorInfoLabel = &H808080
Public Const vbColorTextDarkBlue = &H964901
Public Const vbColorForeFixed = &H606060
Public Const vbColorLightFixed = &HE0E0E0
Public Const vbColorMediumFixed = &HD0D0D0
Public Const vbColorTextLightBlue = &HEBC99B
Public Const vbColorTextBlue = &H8000000D ' &HA55302 '&HB76C00

Public Const vbColorButtonSI = &H644603
Public Const vbColorButtonMouseOver = &H846623
Public Const vbColorButtonNO = &H745613


Public Const vbColorLabelUnabled = &H707070

Public Const vbGreenColor = &H8000000D
Public Const vbTimBlue = &H644603



'------------------------------------------
'       Variabili globali
'------------------------------------------

'Public m_ControlGridFontSize As Double
'Public m_ControlGridRowHeight As Double
'Public m_ControlGridColWidth As Double'

Public Operatore As String
Public PDFPrinter As Boolean

Public SelectProcedura(3) As Boolean

Public MyFormatString As String


Public INDEX_STD As Integer
Public absIndex As Integer

Public bDotForDecimals As Boolean


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



Public Function PadString(strSource As Variant) As String
Dim lPadLen As Integer
Dim PadChar As String
Dim strPad As String
PadChar = ""
'PadString = String(lPadLen, Left(PadChar, 1))
'Mid(PadString, 1, Len(strSource)) = strSource
    strPad = CStr(strSource)
    strPad = Format$(strPad, "#0.000")
   'lPadLen = 12
 ' ' PadString = Left$(String(lPadLen, "") & strSource, lPadLen)
   'PadString = s
   'tring(lPadLen, Left(PadChar, 1))
  ' Mid(PadString, 1, Len(strSource)) = strSource
   
PadString = Format$(strPad, "@@@@@@@@@@@@@@")
End Function

Public Function GetDotForDecimals()


bDotForDecimals = GetSetting(App.Title, "Notation", "bDotForDecimals", False)


End Function


Public Function CheckDot(ByVal Value As String) As Double
On Error GoTo ERR_CHECK
    CheckDot = 0
    If Value <> "" Then
    
        If IsNumeric(Value) Then
            If bDotForDecimals Then
                  
                Value = Replace(Value, ",", ".")
            Else
                Value = Replace(Value, ".", ",")
            
            End If
        Else
            Value = 0
            
        End If
            
        
    Else
        Value = "0"
    End If

ERR_END:
    On Error GoTo 0
    CheckDot = CDbl(Value)
    Exit Function
ERR_CHECK:
    MsgBox "CheckDot Error"
    GoTo ERR_END:
End Function


Public Function BackupDatabase()
Dim bDatabaseCopy As Boolean
Dim DatabaseCopyDate As Date

On Error GoTo ERR_CHECK:

dbPath = GetSetting(App.Title, "ARCHIVIO", "PATH", APP_DATA_FOLDER)

If FileExists(dbPath & dbName) And FileExists(dbPath & dbCodeName) Then
    
    DatabaseCopyDate = FormatDateTime((GetSetting(App.Title, "Database Check", "DatabaseCopyDate", "0.0.00")), vbShortDate)
    bDatabaseCopy = IIf(FormatDateTime(Now(), vbShortDate) = DatabaseCopyDate, False, True)
    
    
    If bDatabaseCopy Then
        Debug.Print dbPath & dbName
        
        FileCopy dbPath & dbName, USER_DOCUMENTI & dbName
        FileCopy dbPath & dbCodeName, USER_DOCUMENTI & dbCodeName
        PopupMessage 2, "FileCopy OK", , , "Database Backup"
        SaveSetting App.Title, "Database Check", "DatabaseCopyDate", FormatDateTime(Now())
    End If

End If

ERR_END:
    On Error GoTo 0
    Exit Function
ERR_CHECK:
    MsgBox err.Description
    GoTo ERR_END:
End Function

Public Function GetDatabaseStartCheck()

On Error GoTo ERR_CHECK:

Dim rc As Boolean

    Call CheckCodeDB

ERR_END:
    On Error GoTo 0
    Exit Function
ERR_CHECK:
    MsgBox err.Description
    GoTo ERR_END:

End Function


Public Function GetDatabaseStartCheckOLD()
Dim bDatabaseCodeClassificationCheck As Boolean

On Error GoTo ERR_CHECK:

Dim rc As Boolean



bDatabaseCodeClassificationCheck = GetSetting(App.Title, "Database Check", "bDatabaseCodeClassificationCheck", False)

If bDatabaseCodeClassificationCheck Then

Else

    rc = CheckCodeClassification
    SaveSetting App.Title, "Database Check", "bDatabaseCodeClassificationCheck", rc

End If


ERR_END:
    On Error GoTo 0
    Exit Function
ERR_CHECK:
    MsgBox err.Description
    GoTo ERR_END:
    
End Function
Public Function CheckCodeClassification() As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim strDescription As String
Dim strCode As String


    On Error GoTo ERR_CHECK:
    
    rc = False
    
    With dbTabCode
        .filter = ""
        .filter = ""
        If .EOF Then
            rc = True
        Else
            .MoveFirst
            For i = 1 To .RecordCount
                strCode = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
                strDescription = IIf(IsNull(Trim(!ProductName)), "", Trim(!ProductName))
                If strCode <> "" Then
                    With dbTabCodeClassification
                        .filter = ""
                        .filter = "Code='" & Replace(strCode, "'", "''") & "'"
                        If .EOF Then
                            .AddNew
                            !Code = strCode
                            !Name = strDescription
                            !DateModified = Now()
                            .Update
                        End If
                    
                    End With
                End If
            
                .MoveNext
            Next
            rc = True
        End If
    End With

ERR_END:
    On Error GoTo 0
    CheckCodeClassification = rc
    Exit Function
ERR_CHECK:
    rc = False
    PopupMessage 2, "CheckCodeClassification Error"
    GoTo ERR_END:
End Function
