Attribute VB_Name = "LotProduction"
Option Explicit


Public Function CheckProductionLot(ByRef Lot As String, ByVal Line As String, ByVal bSetNewLot As Boolean, ByRef uProduction As RecipeForProduction) As Boolean
Dim rc As Boolean
Dim NewLot As String

Dim Tent As Integer
Dim i As Integer

    On Error GoTo ERR_LOT
    rc = True
    '--------------------------------------
    ' ultimo numero possibile
    '--------------------------------------
    Tent = 0
    With dbTabProduction
        .Close
        .Open "SELECT *  FROM TabProduction order by id -1", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
        .filter = ""
        .filter = "Line='" & Line & "'"
        
        If .EOF Then
            NewLot = 1
            NewLot = Format(NewLot, String(4, "0"))
        Else
            .MoveLast
cont:
            NewLot = IIf(IsNull(Trim(!Lot)), "0000", Trim(!Lot))
            
            If IsNumeric(NewLot) And Len(NewLot) = 4 Then
            
                Tent = 0
addnumber:
                '--------------------------------------
                ' incremento il num lot
                '--------------------------------------
                NewLot = NewLot + 1
                '--------------------------------------
                ' se supero LAST_NUMBER torno a capo!
                '--------------------------------------
                If CDbl(LAST_NUMBER) < CDbl(NewLot) Then
                    NewLot = 1
                End If
            Else
                '--------------------------------------
                ' se NON è un numero?
                ' torno indietro di uno....
                ' ho 5 tentativi, poi mollo!
                '--------------------------------------
                .MovePrevious
                Tent = Tent + 1
                If Tent < 5 Then GoTo cont:
                
            End If
            
            
    
            NewLot = Format(NewLot, String(4, "0"))
        
            Tent = Tent + 1
            '--------------------------------------
            ' check se il "nuovo Lotto" esiste
            ' già nel DB
            ' se esiste allora incremento di 1 e
            ' riprovo.... ( ho 5 tentativi, poi
            ' faccio fare all'operatore )
            '--------------------------------------
            If ProdCheckLotAndLine(NewLot, Line) Then
                ' è a tutti gli effetti un lotto nuovo!
                GoTo contCheck
            Else
                ' no : è già presente, aumento il counter e riprovo..
            End If
            
            If Tent < 5 Then GoTo addnumber:
            
        End If
contCheck:
        If Lot = "" Then
            rc = False
             GoTo setNewLot
        End If
    
        If bSetNewLot Then
setNewLot:
            
            Lot = NewLot
            GoTo ERR_END:
        
        End If
        
        .filter = ""
        .filter = "Lot='" & Lot & "' and Line='" & Line & "'"
        If .EOF Then
            rc = True
        Else
        
        
            If uProduction.DateRecipe = IIf(IsNull(Trim(!DataRecipe)), "", Trim(!DataRecipe)) And uProduction.PrepWeek = IIf(IsNull(Trim(!PrepWeek)), "", Trim(!PrepWeek)) And uProduction.numPrepWeek = IIf(IsNull(Trim(!numPrepWeek)), "", Trim(!numPrepWeek)) Then
                '---------------------------------------------------
                ' veloce controllo se ho riaperto una preparation
                '---------------------------------------------------
               
                rc = True
                GoTo ERR_END:
            End If
            
            If UCase(Lot) = IIf(IsNull(Trim(!Lot)), "0000", Trim(!Lot)) Then
                rc = False
            End If
            
            If F_MsgBox.DoShow("This Lot for " & Line & " already exist." & vbCrLf & "Lot available = " & NewLot, "Lot = " & Lot, , "Use New Lot", "Exit") Then
            
                Lot = NewLot
                rc = True
            Else
                Lot = ""
                rc = True
            
            End If


        End If
ERR_END:
        
        CheckProductionLot = rc
        .Close
        .Open "SELECT *  FROM TabProduction order by id ", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
        
    End With
    
    Exit Function
    
ERR_LOT:
    rc = False
    MsgBox err.Description
    Resume Next
    
End Function



Public Function ProdCheckLotAndLine(ByRef Lot As String, ByVal Line As String) As Boolean
Dim rc As Boolean

    rc = False
    
    With dbTabProduction
        .filter = ""
        .filter = "Lot='" & Lot & "' and Line='" & Line & "'"
        If .EOF Then
            rc = True
        Else
        
            rc = False
        End If
    End With
    
    ProdCheckLotAndLine = rc
    
End Function


