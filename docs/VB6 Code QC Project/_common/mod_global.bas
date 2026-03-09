Attribute VB_Name = "mod_global"
Option Explicit


'------------------------------------------
'       Variabili Globali
'------------------------------------------

' per spostare la FORM

Public FrmMove As Boolean
Public DragX, DragY   As Single

'------------------------------------------

Public bFullScreen              As Boolean
Public PictureMaxScreen         As String
Public MinNumberSelecterPerc    As Double
Public SettingName              As String
Public UsedFileName             As String
Public PdfFileName              As String
Public bSearchClosedLot        As Boolean



'------------------------------------------
'       Costanti
'------------------------------------------

Public Const USER_ESTENSIONE = ".qc"
Public Const USER_ESTENSIONE_RICHIESTE = ".rp"
Public Const ProjectName = "ChemicalQC"
Public Const MaxTipoLetture = 10

'------------------------------------------
'       colori standard
'------------------------------------------
Public Const vbColorAzzurrino = &HFFE1C1
Public Const vbColorRosaTabella = &HC7B8FE
Public Const vbColorRed = &HC0&
Public Const vbColorTextDarkBlue = &H964901
Public Const vbColorgotFocus = &HF0F0F0
Public Const vbColorUnabled = &HC0C0C0
Public Const vbColorDarkUnabled = &H202020
Public Const vbColorOrange = &H80FF&
Public Const vbColorDarkFont = &H404040
Public Const vbColorGreen = &H4000&
Public Const vbColorInfoLabel = &H808080
Public Const vbColorForeFixed = &H606060
Public Const vbColorLightFixed = &HE0E0E0
Public Const vbColorMediumFixed = &HD0D0D0
Public Const vbColorTextLightBlue = &HEBC99B
Public Const vbColorTextBlue = &H8000000D ' &HA55302 '&HB76C00

Public Const vbColorLabelUnabled = &H707070

Public Const vbGreenColor = &H8000000D
Public Const vbTimBlue = &H644603

Public Const vbColorBlueProgram = &H964901

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

'------------------------------------------
'
'       Tipi globali
'
'-------------------------------------------


Public Type TypeOperatore
    Name As String
    IndexPrivilege As Integer
End Type

Public Type Postazione
   Num As Integer
   Descrizione As String
   Operatore As String
   Enabled As Boolean
   Department As String
End Type

Public Type HannaCode
   FGCode           As String
   Code             As String
   ProductName      As String
   Description      As String
   Recipe           As String
   Line             As String
   QCMethod         As String
   MeterFamily1     As String
   MeterFamily2     As String
   ParameterMethod  As String
   ParameterUnit    As String
   MeasurementUnit  As String
   RangeMin         As String
   RangeMax         As String
   Decimal          As String
   WeightValue      As String
   WeightMin        As String
   WeightMax        As String
   Certified        As Boolean
   
End Type

Public Type TestReading
    Value           As String
    Meter           As Integer
    Test            As Integer
    bSelectedValue  As Boolean
    
    STDValueAvg     As String
    STDAbs          As String
    
End Type

Public Type ptCheck

    Time                As String
    Enabled             As Boolean
    Meter(4)            As Integer
    Readings(80)       As TestReading
    MaxReadings         As Integer
    SelReadings         As Integer
    NumTest             As Integer
    SelTest             As Integer
    TotalMean           As Double
    SelecMean           As Double
    
End Type

Public Type DataControllo
    media               As Double
    devst               As Double
    LC                  As Double
    Tolerance           As Double
    STDRef              As Double
    STDNumber           As String
    STDMin              As String
    STDMax              As String
    NumReadings         As Integer
    MeasurementUnit     As String
    s3                  As Double
    s                   As Double
    s2                  As Double
    Note                As String
    OutOfRangeData      As Integer
    OutOfRangeDataPerc  As Double
    Operator            As String
End Type

Public Type QCDefault

    Code                As String
    Description         As String
    Expiration          As String

End Type



Public Type ptChemicalQC
    Lot                 As String
    date                As String
    Exp                 As String
    Note                As String
    QCOperator          As String
    HannaCode           As HannaCode
    QCReagentA(6)      As QCDefault
    QCReagentB(6)      As QCDefault
    Meter(4)            As QCDefault
    ph(2)               As QCDefault
    PrepWeek            As String
    PrepOperator        As String
    ProdFirst           As String
    ProdLast            As String
    ProdMachine         As String
    OldLotA             As String
    OldLotAExpiration   As String
    OldLotB             As String
    OldLotBExpiration   As String
    Department          As String
    RegistrationBook    As String
    QCType              As String
    DataControllo(15)    As DataControllo
    FileName            As String
    NLot                As String
    nSETTING            As String
    numHeads            As Integer
    type                As Integer
    desc_type           As String
    Enabled             As Boolean
    STDtest(10)         As ptCheck
    NumReadings         As Integer
    AllReadings         As Integer
End Type


Public MyChemicalQC         As ptChemicalQC
Public MyGraphicCheck       As ptChemicalQC

Public MyChemicalQCClean    As ptChemicalQC



Public MyOperatore          As TypeOperatore
Public MyOperatoreClean     As TypeOperatore
Public MyHannaCode          As HannaCode
Public MyHannaCodeClean     As HannaCode

Public MyImportHannaCode As HannaCode




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


