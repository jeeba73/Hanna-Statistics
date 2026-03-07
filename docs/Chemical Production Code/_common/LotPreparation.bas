Attribute VB_Name = "LotPreparation"
Option Explicit

Public Const LAST_NUMBER As String = "9999"
Public iLotNumberType As Integer


Public Function CheckLotNumberType()

iLotNumberType = GetSetting(App.Title, "LotNumber", "iLotNumberType", 0)

End Function



Public Function CheckPreparationLot(ByRef Lot As String, ByVal Line As String, ByVal bSetNewLot As Boolean, ByRef uPreparation As RecipeForProduction) As Boolean
Dim rc As Boolean
Dim NewLot As String
Dim PrepDate As String
Dim numPrepWeek As String

Dim Tent As Integer
Dim i As Integer

    On Error GoTo ERR_LOT
    rc = True
    '--------------------------------------
    ' ultimo numero possibile
    '--------------------------------------
    

    With dbTabPreparation
    
    
            If Not IsNumeric(Lot) Then
                Lot = ""
                GoTo NewLot
            End If
            
            Lot = LotNumberControl(Lot)
            PrepDate = uPreparation.PreparationDate
            numPrepWeek = uPreparation.numPrepWeek
            Tent = 0
            
            Lot = Format$(Lot, "0000")
            
            
NewLot:
        .Close
        .Open "SELECT *  FROM TabPreparation order by id -1", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
        .filter = ""
        .filter = "Line='" & Line & "' and Lot <> NULL "
        
        If .EOF Then
            NewLot = 1
            NewLot = LotNumberControl(NewLot)
            NewLot = Format(NewLot, String(4, "0"))
           
        Else
            If .RecordCount > 1 Then
                .MoveLast
                .MovePrevious
            Else
                .MoveFirst
            End If
cont:
            NewLot = IIf(IsNull(Trim(!Lot)) Or Trim(!Lot) = "", "0000", Trim(!Lot))
            
            If IsNumeric(NewLot) And Len(NewLot) = 4 Then
            
                Tent = 0
addnumber:
                '--------------------------------------
                ' incremento il num lot
                '--------------------------------------
                NewLot = NewLot + 1
                
                NewLot = LotNumberControl(NewLot)
                
                
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
            
            
               NewLot = Format$(NewLot, "0000")
            'NewLot = Format(NewLot, String(4, "0"))
        
            Tent = Tent + 1
            '--------------------------------------
            ' check se il "nuovo Lotto" esiste
            ' già nel DB
            ' se esiste allora incremento di 1 e
            ' riprovo.... ( ho 5 tentativi, poi
            ' faccio fare all'operatore )
            '--------------------------------------
            If CheckLotAndLine(NewLot, Line) Then
                ' è a tutti gli effetti un lotto nuovo!
                GoTo contCheck:
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
        .filter = "Lot='" & Lot & "' and Line='" & Line & "'" ' and PrepDate<>'" & PrepDate & "' and numPrepWeek<>'" & numPrepWeek & "'"
        If .EOF Then
            rc = True
        Else
        
            'If !PrepDate = PrepDate And !numPrepWeek = numPrepWeek Then

               ' GoTo ERR_END
            'End If
            If uPreparation.Recipes(1).Code = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe)) And uPreparation.DateRecipe = IIf(IsNull(Trim(!DataRecipe)), "", Trim(!DataRecipe)) And uPreparation.PrepWeek = IIf(IsNull(Trim(!PrepWeek)), "", Trim(!PrepWeek)) And uPreparation.numPrepWeek = IIf(IsNull(Trim(!numPrepWeek)), "", Trim(!numPrepWeek)) Then
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
        
        Lot = Format$(Lot, "0000")
        CheckPreparationLot = rc
        .Close
        .Open "SELECT *  FROM TabPreparation order by id ", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
        
    End With
    
    Exit Function
    
ERR_LOT:
    rc = False
    MsgBox err.Description
    Resume Next
    
End Function

Public Function LotNumberControl(ByVal Lot As Long) As Long

LotNumberControl = Lot

Select Case iLotNumberType
    Case 0
    Case 1 ' pari
        If Lot / 2 = Int(Lot / 2) Then
        Else
            LotNumberControl = Lot + 1
            'PopupMessage 2, "Lot Number must be a Even Number" & vbCrLf & "New Lot Number = " & LotNumberControl, , , "EVEN LOT NUMBER"
        End If
        
    Case 2 ' dispari
        If Lot / 2 <> Int(Lot / 2) Then
        Else
            LotNumberControl = Lot + 1
            'PopupMessage 2, "Lot Number must be a Odd Number" & vbCrLf & "New Lot Number = " & LotNumberControl, , , "ODD LOT NUMBER"
        End If
End Select
                
End Function



Public Function CheckLotAndLine(ByRef Lot As String, ByVal Line As String) As Boolean
Dim rc As Boolean

    rc = False
    
    With dbTabPreparation
        .filter = ""
        .filter = "Lot='" & Lot & "' and Line='" & Line & "'"
        If .EOF Then
            rc = True
        Else
        
            rc = False
        End If
    End With
    
    CheckLotAndLine = rc
    
End Function

