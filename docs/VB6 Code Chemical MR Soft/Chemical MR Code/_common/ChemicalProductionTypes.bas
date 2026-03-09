Attribute VB_Name = "ChemicalMRTypes"
Option Explicit


'------------------------------------------
'
'       Tipi globali
'
'-------------------------------------------

' Standard

Public Type STD
    MRCode            As String
    NUMBER            As Integer
    Value             As Double
    TheoreticalWeight As Double
    RealWeight        As Double
    ActualWeight      As Double
    Unit              As String
    bOK               As Boolean
    bChanged          As Boolean
    Variance          As Double
    VariancePerc      As Double
    Note              As String
    STD_ID            As String
End Type


' barcode type

Public Type Barcode
    Code             As String
    Date             As String
    Lot              As String
    Bottle           As String
End Type


Public Type TypeOperatore 'ok
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


Public Type Pipette 'ok
    Equipment           As String
    VolumeAdjustment    As String
    Characteristic      As String
    VolMin              As String
    VolMax              As String
    Unit                As String
    Decimal             As String
    GradResolution      As String
    Acc                 As String
    AccMin              As String
    AccMax              As String
    ImmersionDeph       As String
    WaitTime            As String
End Type



Public Type FrasiH 'ok
    Code             As String
    PhyHazStatement  As String
    HazardCategory   As String
    Precaution       As String
    SafetyEquipments As String
    Pictogram        As String
End Type

Public Type ProductClassification 'ok
    Code            As String
    Name            As String
    Cas             As String
    Index           As String
    Cee             As String
    Recipe          As String
    Phrases         As String
    CounterFrasi    As Integer
    FrasiH()        As FrasiH
    
End Type


Public Type MRType
    Code                    As String
    Description             As String
    Supplier                As String
    
    
  
    MRPurity                As Double
    MRValue                 As Double

    MNP                     As String
    PhysicalState           As String
    Density                 As String
    Unit                    As String
    Parameter               As String
    FWParameter             As Double
    StorageT                As String
    MinQty                  As String
    Modified                As Date

    STOCK_QTY               As Double
    STOCK_UNIT              As String
    Rev                     As String
    
    ReductionExpDays        As String
    
    Classification          As String
    
    bMassa                  As Boolean
    
    Location                As String
    
End Type

Public Type WareHouseEntry
    ID                      As Long
    MRCode                  As String
    Description             As String
    Bottle()                As String
    EntryBottle             As String
    NumberBottle            As Integer
    Lot                     As String
    Density                 As String
    Purity                  As String
    MRValueConcentration    As String
    Unit                    As String
    Parameter               As String
    FWParameter             As String
    Location                As String
    StockQTY                As String
    stockUnit               As String
    ArrivedTime             As String
    Open                    As String
    Finished                As String
    SupplierEXP             As String
    MREXP                   As String
    MNP                     As String
    Status                  As String
    Note                    As String
    Operator                As String
    bBarcode()              As Boolean
    BarcodeText()           As String
    U                       As String
    LastLetter              As String
    PreparationID           As String
    
End Type

Public Type HannaCode 'ok



    RangeMin            As String ' solo per import excel
    RangeMax            As String ' solo per import excel
    ProductName         As String
        
    ID                  As Long
    Code                As String
    Description         As String
    Recipe              As String
    
    MR                  As MRType

    ParameterMethod     As String
    Hannaformula    As String
    FWHannaParameter  As Double
    ConcHannaParameter  As Double
    MeasurementUnit     As String
    Decimal             As Double
    
    STDMR2              As String
    MS1val              As String
    MS1vol              As String
    MS2Dil              As String
    MS2vol              As String
    MSEXP               As String
    STDMatrix           As String
    STDVolume           As String
    STDUnit             As String
    STDExp              As String
    STDExpDate          As String
    STDNote             As String
    STDStorage          As String
    UnitMR              As String
    QtyToProduce        As String
    STD()               As STD
    STDcount            As Integer
    STDType             As Integer
End Type

' mother solution


Public Type MotherSolution
    ID                As Long
    Code              As String
    HannaCode         As String
    DataPrep          As Date
    HourPrep          As String
    ExpDays           As String ' se nullo = 10
    DataExp           As Date
    WeekPrep          As String
    DataMS            As Date
    Operator          As String
    QtyLeft           As Double
    QtyUsed           As Double
    QtyProduced       As Double
    Unit              As String
    bClosed           As Boolean
    Note              As String
    MsType            As Integer
    PreparationID     As Long
    Bottle            As WareHouseEntry
    
End Type


Public Type PrepAcquisition

    AcquisitionTime      As Date
    Index                As Integer
    ID                   As Long
    
    Code                 As String
    HannaCode            As String
    Bottle               As String
    MRLot                As String
    DatePrep             As Date
    HourPrep             As String
    WeekPrep             As String
    FileName             As String
    
    PrepBarcode          As Barcode
    ActualWeight         As Double ' quello che l'operatore utilizza
    bFromBarcode         As Boolean
    
    Note                 As String
  
    Operator             As String
    
    bDeleted             As Boolean
    
    PreparationID        As String
    MsType               As Integer
    STDNumber            As String
    STDValue             As String
    STDQty               As String
    STDUnit              As String
    MotherSolutionDate   As String
    LeftInBottle         As String
    CodicePipetta        As String
    PipettaType          As String
    ScaleID              As String
    GlasswareID          As String
    bManuale             As Boolean
    
    MNP                  As String
    ExpMR                As String
    
End Type



Public Type RecipeType 'ok
    Type                        As String
    ID                          As Long
    Code                        As String
    Description                 As String
    HannaCode                   As HannaCode

    ' preparation
    ActualWeight                As Double
    TotalWeight                 As Double
    bRecalculation              As Boolean
    bUmMassa                    As Boolean
    MotherSolution              As MotherSolution
    STD()                       As STD
    STDcount                    As Integer
    STDUnit                     As String
    ' preparation
    AcquisitionCount    As Integer
    Acquisitions()      As PrepAcquisition
    
    bManual                     As Boolean
End Type

Public Type MotherSolSpec

    DilConc       As Double
    Volume        As Double
    Qty           As Double
    Value         As Double
    Unit          As String
    Exp           As Integer
End Type

Public Type RecipeForProduction
    Recipe            As RecipeType
    HannaCode          As HannaCode
    DataPrep          As Date
    HourPrep            As String
    MRCode          As String
    ExpDate As Date
    
    PrepWeek          As String
    Operator          As String
    QtyToProduce      As Double
    QtyProduced       As Double
    Unit              As String
    bClosed           As Boolean
    CloseDate         As String
    Note              As String
    FileName          As String
    MotherSol         As MotherSolution
    MS                As MotherSolSpec
    MsType            As Integer
    ID                As Long
    AcquisitionCount    As Integer
    bPesatoTuttiComponenti  As Boolean
    bCorrection             As Boolean
    
    bManual             As Boolean
    ManualValue         As Double
    ManualUnit          As String
    STD_Manual_ID       As String
    

End Type



Public RecipeClean As RecipeType
Public uPreparation As RecipeForProduction
Public uPreparationClean As RecipeForProduction



Public DataPipette() As Pipette
Public DataPipetteClean() As Pipette

Public MyImportHannaCode As HannaCode
Public MRCleanArray() As MRType
Public MyOperatore          As TypeOperatore
Public MyOperatoreClean     As TypeOperatore
Public MyHannaCode          As HannaCode
Public MyHannaCodeClean     As HannaCode

Public MRTypeClean As MRType
Public MotherSolutionClean As MotherSolution
Public MyWareHouseEntryCleanArray() As WareHouseEntry
Public MyWareHouseEntryClean As WareHouseEntry
Public UserBarcodeClean As Barcode
Public MyFrasiH() As FrasiH
Public MyFrasiHClean() As FrasiH
Public MyImportProductClassification() As ProductClassification
Public MyImportProductClassificationClean() As ProductClassification



Public Function InvUm(unita) As Double
    
    Select Case LCase(unita)
        Case "µg"
            InvUm = 1000000
        Case "mg"
            InvUm = 1000
        Case "kg"
            InvUm = 0.001
        Case "l"
            InvUm = 0.001
        Case "t"
            InvUm = 0.000001
        Case Else
            InvUm = 1
    End Select
End Function


Public Function UmMS(ByVal unita As String) As Double

Dim splitMe As Variant
Dim str As String

    str = unita
    splitMe = Split(str, "/")

    unita = splitMe(0)

    Select Case LCase(unita)
        Case "µg"
            UmMS = 1000
        Case "mg"
            UmMS = 1
        Case "g"
            UmMS = 0.001
    End Select
    
    
    
End Function

Public Function Um(ByVal unita As String) As Double

Dim splitMe As Variant
Dim str As String

    str = unita
    splitMe = Split(str, "/")


    unita = splitMe(0)



    Select Case LCase(unita)
        Case "µg"
            Um = 0.000001
        Case "mg", "ul"
            Um = 0.001
        Case "g", "ml"
            Um = 1
        Case "kg", "l"
            Um = 1000
        Case "t"
            Um = 1000000
        Case Else
            Um = 1
    End Select
End Function

Public Function SetbUmMassa(ByVal unita As String) As Boolean
Select Case LCase(unita)

        Case "kg"
            SetbUmMassa = True
        Case "g"
            SetbUmMassa = True
        Case "mg"
            SetbUmMassa = True
         Case "l"
            SetbUmMassa = False
        Case "ml"
            SetbUmMassa = False
    End Select
    
End Function
Public Function SetUmVolume(unita) As String
    Select Case LCase(unita)

        Case "kg"
            SetUmVolume = "L"
        Case "g"
            SetUmVolume = "mL"
        Case "mg"
            SetUmVolume = "mL"
         Case "l"
            SetUmVolume = "Kg"
        Case "ml"
            SetUmVolume = "g"
    End Select
End Function

Public Function SetUmComponent(unita) As String
    Select Case LCase(unita)

        Case "kg"
            SetUmComponent = "g"
        Case "g"
            SetUmComponent = "g"
        Case "mg"
            SetUmComponent = "mg"
         Case "l"
            SetUmComponent = "g"
        Case "ml"
            SetUmComponent = "g"
    End Select
End Function

Public Function iVirgola(ByVal PesoInGrammi As Double) As Integer


'CIFRE DECIMALI: 3 SE < 10, 2 SE COMPRESE TRA 10 E 100 - 1 SE COMPRESE TRA 100 E 1000 - 0 SE > 1000
' il valore può essere o g o mL



    Select Case PesoInGrammi
        Case Is < 10
            iVirgola = 3
        Case 10 To 100
            iVirgola = 2
        Case 100.01 To 1000
            iVirgola = 1
        Case Is > 1000
            iVirgola = 0
    End Select



End Function

Public Function SetUmCombo(ByRef cmb As ComboBox, ByVal bMassa As Boolean, ByVal bTutti As Boolean)
    
    If bTutti Then bMassa = True
    
    With cmb
        .Clear
        If bMassa Then
            
            .AddItem "µg"
            .AddItem "mg"
            .AddItem "g"
            .AddItem "kg"
            If bTutti Then GoTo Liquid
        Else
Liquid:
            .AddItem "mL"
            .AddItem "L"
        End If
        .ListIndex = 0
    End With

End Function
