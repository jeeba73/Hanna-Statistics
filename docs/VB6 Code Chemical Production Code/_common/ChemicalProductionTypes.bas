Attribute VB_Name = "ChemicalProductionTypes"
Option Explicit


'------------------------------------------
'
'       Tipi globali
'
'-------------------------------------------

Public Type QCType
    Status      As String
    QcClosed    As Boolean
    Operator    As String
    Date        As Date
    Note        As String
    Index       As Integer
    RecipeCode  As String
    ID    As Long
    SettingName As String
    Registration    As String
    QCOperator      As String
    Correction      As String
    CorrectionDate  As String
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

Public Type ProdWay 'ok

    Recipe          As String
    Production      As String
    Head            As String
    Speed           As Integer
    Line            As String
    EstTimeH        As String
    EsttimeD        As String
    
End Type


Public Type RawMaterial 'ok

    Code                    As String
    Description             As String
    Cas                     As String
    ChemicalReactionLiquid  As String
    Classification          As String
    Pictograms              As String
    Um                      As String
    ManufacturerName        As String
    ManufacturerCode        As String
    Location                As String
    SpecifiedLocation       As String
    bMix                    As Boolean
    CriticalRM              As String
    Density                 As Double
    
End Type
Public Type RMRecipeForProduction 'ok

   
    ChemicalReactionLiquid  As String
    Classification          As String
    Pictograms              As String
    ManufacturerName        As String
    ManufacturerCode        As String
    Location                As String
    SpecifiedLocation       As String
    CriticalRM              As String
    QtyToProduce            As Double
    UmQty                   As String
    Tollerance              As Double
    
End Type

Public Type Totals

    Recipe                  As String
    Description             As String
    TotalWeighKg            As String
    TotalWeighL             As String
    TotalMultiple           As String
    CkMin                   As Boolean
    CkMax                   As Boolean
    Min                     As String
    Max                     As String
    Minpcs                  As String
    bMix                    As Boolean
    Multiple                As String


End Type
Public Type ProdAcquisition
    Index               As Integer
    Code                As String
    LotNumber           As String
    DateProd            As String
    Machine             As String
    WeekProd            As String
    Note                As String
    Operator            As String
    AcquisitionTime     As Date
    ID                  As Long
    bDeleted            As Boolean
    QtyProduced         As Double
    Mix1Lot             As String
    Mix2Lot             As String
    ExpDate             As String
End Type

Public Type HannaCode 'ok
    ID                      As Long
    Index                   As Integer
    bHide                   As Boolean
    Code                    As String
    Line                    As String
    STD                     As String
  '  LoadInPrint             As String
    ProductName             As String
    Recipe                  As String
    Mix1                    As String
    Mix2                    As String
    Exp                     As String
    ExpDate                 As String
    Um                      As String
    Qty                     As String
    QtyToProduce            As String
    QtyProduced             As String
    Density                 As String
    MinQty                  As String
    MaxQty                  As String
    UncertantlyFromCoA      As String
    Procedure               As String
    ProcedureRev            As String
    LastLot                 As String
    LotNumber               As String
    
    DateProd                As String
    Machine                 As String
    WeekProd                As String
    ' production
    AcquisitionCount        As Integer
    Acquisitions()          As ProdAcquisition
    
    RangeMin            As String ' solo per import excel
    RangeMax            As String ' solo per import excel
        
End Type

' recipe for production types

Public Type MaterialRequisition
    NUMBER              As String
    Operator            As String
    Today               As Date
    Reason              As String
    ProductionLine      As String
End Type


' preparation types
' barcode type
Public Type Barcode
    Code                    As String
    Cas                     As String
    ChemicalName            As String
    Manufacturer            As String
    ManufacturerCode        As String
    ManufacturerLot         As String
    DeliveryDate            As String
    QtyDelivered            As String
    Package                 As String
    WeekDelPackageNumber    As String
    
End Type


Public Type PrepAcquisition
    Index               As Integer
    PrepBarcode         As Barcode
    ActualWeight        As Double
    bRecalculation      As Boolean
    bRecipeComponent    As Boolean
    bFromBarcode        As Boolean
    Note                As String
    Operator            As String
    AcquisitionTime     As Date  '+
    ID                  As Long
    bDeleted            As Boolean
    ExpDate             As String
End Type


' all types

Public Type RmxRecipe 'ok
    ID                  As Long
    bDeleted            As Boolean
    RecipeCode          As String
    CHCode              As String
    Description         As String
    Cas                 As String
    Qty                 As Double
    Um                  As String
    Perc                As Double
    RealPerc            As Double
    Note                As String
    bMix                As Boolean
    TheoreticalWeight   As Double
    ActualTheoreticalWeight As Double
    Variance            As Double
    VariancePerc        As Double
    UmTheoreticalWeight As String
    TotalWeightKg       As Double
    TotalWeightL        As Double
    TotalMultiple       As Double
    ManufacturerLot     As String
    Specifications      As RMRecipeForProduction
    
    Density             As Double
    bUmMassa            As Boolean
    MaxQty              As Double
    UmMax               As String
    MinQty              As Double
    Multiple            As Double
    MultipleInCell      As String
    MultipleToProduce   As Double
    MultipleMassa       As Double
    UmMultiple          As String
    MinQty2             As Double
    UmMinQty            As String
    
    
  
    RealWeight          As Double
    bCorrection         As Boolean
    TolerancePerc       As Double
    bAddedInPreparation As Boolean
    
    CriticalRM          As String
End Type



Public Type RecipeType 'ok
    
    ID                          As Long
    bHaveMixes                  As Boolean
    bIsMix                      As Boolean
    bTestLot                    As Boolean
    PreparationLotMix           As String
    Code                        As String
    Machine                     As String
    Description                 As String
    Line                        As String
    Procedure                   As String
    ProcedureDate               As String
    Rev                         As String
    NoteRev                     As String
    RevDate                     As String
    Classification              As String
    Exp                         As String
    ExpDate                     As String
    Density                     As Double
    bUmMassa                    As Boolean
    MaxQty                      As Double
    UmMax                       As String
    MinQty                      As Double
    Multiple                    As Double
    MultipleToProduce           As Double
    MultipleMassa               As Double
    UmMultiple                  As String
    TotalRecipe                 As String
    MinQty2                     As Double
    UmMinQty                    As String
    TotalWeightKg               As Double
    UmTotalWeightKg             As String
    TotalMultiple               As Double
    TotalWeightL                As Double
    UmTotalWeightL              As String
    Component()                 As RmxRecipe
    ComponentCount              As Integer
    RmxRecipe()                 As RmxRecipe
    RmxRecipeCount              As Integer
    HannaCodes()                As HannaCode
    HannaCodesCount             As Integer
    bUpdated                    As Boolean
    bNoPreparation              As Boolean
    bHide                       As Boolean
    ProductionWay               As ProdWay
    MaterialRequisition()       As MaterialRequisition
    MaterialRequisitionCount    As Integer
    Mix                         As String
    
    Cas                         As String
    Location                    As String
    SpecifiedLocation           As String
    
    ' preparation
    ActualWeight                As Double
    ActualWeightUm              As String
    bRecalculation              As Boolean
    RecipeComponentCount        As Integer
    
    ' preparation
    AcquisitionCount    As Integer
    Acquisitions()      As PrepAcquisition
    
End Type


Public Type RecipeForProduction
    bTestLot            As Boolean
    bAllMixes           As Boolean
    Recipes()           As RecipeType
    HannaCodes()        As HannaCode
    RecipeCount         As Integer
    HannaCodesCount     As Integer
    bOpen               As Boolean
    bSaved              As Boolean
    RecipeBy            As String
    numPrepWeek         As String
    PrepWeek            As String
    PlannedPrepWeek     As String
    Note                As String
    PlanningReference   As String
    DateRecipe          As Date
    PreparationDate     As String
    PreparationLot      As String
    ExpDate             As String
    OperatorPrep        As String
    OperatorProd        As String
    OperatorRfP         As String
    TotalCount          As Integer
    TotalGrid()         As Totals
    PackagingCount      As Integer
    Packaging()         As ProdWay
    fileNameRecForProd  As String
    AcquisitionCount    As Integer
    bPesatoTuttiComponenti  As Boolean
    bCorrection             As Boolean
    QCCount             As Integer
    QCStatus()          As QCType
    WeekProd            As String
    
    ProductionID        As Long
    

End Type


Public Type PreparationType
    Recipe              As RecipeType
    InexRecipe          As Integer
    bOpen               As Boolean
    bSaved              As Boolean
    Operator            As String
    numPrepWeek            As String
    PlannedPrepWeek     As String
    Note                As String
    PlanningReference   As String
    DateRecipe          As Date
    QtyToProduce        As Double
    QtyProduced         As Double
    QCCount             As Integer
    QCStatus()          As QCType
    FileName            As String
    ID                  As Long
End Type


Public RecipeQCClean As QCType

Public RecipeClean As RecipeType
Public uRecipeForProduction As RecipeForProduction
Public uRecipeForProductionClean As RecipeForProduction

Public uPreparation As RecipeForProduction
Public uPreparationClean As RecipeForProduction

Public MyRecipe As RecipeType
Public RmxRecipeClean As RmxRecipe
Public RmxRecipe As RmxRecipe

Public MyRawMaterialClean() As RawMaterial
Public MyRawMaterial() As RawMaterial
Public ExcelProductionWay() As ProdWay
Public ProductionWayClean() As ProdWay
Public MyFrasiH() As FrasiH
Public MyFrasiHClean() As FrasiH
Public MyImportProductClassification() As ProductClassification
Public MyImportProductClassificationClean() As ProductClassification

Public MyImportHannaCode() As HannaCode

Public MyOperatore          As TypeOperatore
Public MyOperatoreClean     As TypeOperatore
Public MyHannaCode          As HannaCode
Public MyHannaCodeClean     As HannaCode


Public Function InvUm(unita) As Double
    
    Select Case LCase(unita)
        Case "ug"
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
Public Function Um(unita) As Double
    Select Case LCase(unita)
        Case "ug"
            Um = 0.000001
        Case "mg"
            Um = 0.001
        Case "mL"
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

Public Function IsLocked(ByRef iRecipe() As RecipeType) As Boolean

 On Error Resume Next
    ReDim Preserve iRecipe(LBound(iRecipe) To UBound(iRecipe))
    IsLocked = err = 10

End Function
