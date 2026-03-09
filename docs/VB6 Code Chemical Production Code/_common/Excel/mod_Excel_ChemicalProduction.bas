Attribute VB_Name = "mod_Excel_ChemicalProduction"
Option Explicit









Public Function EsportaPreparationExcel(ByVal FileName As String, ByVal sString As String, ByRef iPreparation As RecipeForProduction) As Boolean
    Dim rc As Boolean
    On Error GoTo ERR_IMP
    rc = True
   ' MsgBox USER_DESKTOP & "\" & "Backup VerPeriodica.xls"
   
   SettingName = FileName
   
   If SettingName = "" Then Exit Function
   
    
        If CreateExcel(False) Then
            NewExcelWorksheet (sString)
            If CopyChemicalProductionData(SettingName, iPreparation) Then
                Call SaveExcel(sString)
                Call CloseExcel
                PopupMessage 2, "Excel file correctly generated..."
            Else
                rc = False
            End If
        Else
            rc = False
        End If
ERR_END:
    On Error GoTo 0
    EsportaPreparationExcel = rc
    Exit Function
ERR_IMP:
    rc = False
    MsgBox err.Description
    Resume ERR_END
End Function


Public Function CopyChemicalProductionData(ByVal SettingName As String, ByRef iPreparation As RecipeForProduction) As Boolean
Dim rc As Boolean
    On Error GoTo ERR_COPY

    rc = True
ERR_END:
    On Error GoTo 0
    CopyChemicalProductionData = rc
    Exit Function
ERR_COPY:
    rc = False
    MsgBox err.Description
    Resume Next
End Function

