Attribute VB_Name = "hProduction_02_Functions"
Option Explicit

Public Function AggiornaTabSTDPreparation(ByRef STDPreparationID As Long, ByRef uSTDPreparation As RecipeForSTDPreparation)
Dim strRecipes As String
With dbTabSTDPreparation
    .filter = ""
    
    If STDPreparationID = 0 Then
        .AddNew
        GoTo cont:
    End If
    
    .filter = "ID='" & STDPreparationID & "'"
    If .EOF Then
        .AddNew
    End If
cont:
       !Line = uSTDPreparation.HannaCodes(1).Line
       !HannaCode = GetStrHannaCode(uSTDPreparation.HannaCodes(), strRecipes)
       !Recipe = Left$(IIf(IsNull(Trim(!Recipe)), strRecipes, !Recipe), 250)
       !PrepDate = IIf((uSTDPreparation.PreparationDate = "0.00.00") Or IsNull(uSTDPreparation.PreparationDate), uSTDPreparation.DateRecipe, uSTDPreparation.PreparationDate)
       !ExpDate = uSTDPreparation.ExpDate
       !PrepWeek = uSTDPreparation.PrepWeek
       !RecipeWeek = IIf(IsNull(!RecipeWeek), uSTDPreparation.PrepWeek, !RecipeWeek)
       !PlannedPreparation = IIf(IsNull(!PlannedPreparation), uSTDPreparation.PlannedPrepWeek, !PlannedPreparation)
       !numPrepWeek = uSTDPreparation.numPrepWeek
       !PlanningReference = uSTDPreparation.PlanningReference
       !DataRecipe = IIf((uSTDPreparation.DateRecipe = "0.00.00") Or IsNull(uSTDPreparation.DateRecipe), uSTDPreparation.PreparationDate, uSTDPreparation.DateRecipe)
       !OperatorRfP = uSTDPreparation.RecipeBy
       !Note = uSTDPreparation.Note
       !FileName = uSTDPreparation.fileNameRecForProd
       !startDate = IIf(IsNull(!startDate), Now(), !startDate)
       !RfpID = 0
        .Update
        
        STDPreparationID = !ID
    



End With

End Function




Public Function ViewHannaCodeInSTDPreparation(ByRef iHannaCodes() As HannaCode, ByVal Grid1 As Grid, ByVal bView As Boolean) As Boolean
Dim i As Integer
Dim bIsMix As Boolean
Dim Produced As String

    If bView Then
    
        With Grid1
            If .Rows < 1 Then Exit Function
            .AutoRedraw = False
            For i = 1 To .Rows - 1
                .RowHeight(i) = 25
                iHannaCodes(i).bHide = False
            Next
            .Refresh
            .AutoRedraw = True
        End With
 
    Else
        
        Dim strValue As String
        Dim dValue As Double
        
        Dim strValueProduced As String
        Dim dValueProduced As Double
        
        With Grid1
            If .Rows < 1 Then Exit Function
            
            .AutoRedraw = False
                For i = 1 To .Rows - 1
                    strValue = .Cell(i, 6).Text
                    strValue = Replace(LCase(strValue), "kg", "")
                    strValue = Replace(LCase(strValue), "l", "")
                    
                    strValueProduced = .Cell(i, 7).Text
                    If strValueProduced = "" Then strValueProduced = "0"
                    If strValue <> "" Then
                        dValue = CDbl(strValue)
                        If dValue > 0 Then
                        ElseIf dValue = 0 And CDbl(strValueProduced) > 0 Then
                        Else
                            .RowHeight(i) = 0
                            
                            iHannaCodes(i).bHide = True
                        End If
                    Else
                    
                        If CDbl(strValueProduced) > 0 Then
                        Else
                            .RowHeight(i) = 0
                            iHannaCodes(i).bHide = True
                        End If
                    End If
                Next
            .Refresh
            .AutoRedraw = True
        End With
    End If
End Function



Public Function AddCodeInSTDPreparationGrid(ByVal Grid1 As Grid, ByVal HannaCode As String, ByRef uHannaCode() As HannaCode) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim Recipe As String
Dim Mix1 As String
Dim Mix2 As String
Dim t As Integer

            
On Error GoTo ERR_ADD:

    rc = True
    
    HannaCode = Trim(HannaCode)
     
        With Grid1
        
            If .Rows > 1 Then
                For i = 1 To .Rows - 1
                    
                        If Trim(LCase(.Cell(i, 1).Text)) = Trim(LCase(HannaCode)) Then
                            
                            If F_MsgBox.DoShow("Warning : Hanna Code already in Table!", HannaCode, , "Add Again", "Don't") Then
                                Exit For
                            Else
                                rc = False
                                GoTo ERR_END
                            End If
                  
                    Else
                    
                    End If
                Next
            End If
        End With
        
            With dbTabCode
                .filter = ""
                .filter = "Code='" & HannaCode & "'"
                If .EOF Then
            
                Else
                    
                    Recipe = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
                    Mix1 = IIf(IsNull(Trim(!Mix1)), "", Trim(!Mix1))
                    Mix2 = IIf(IsNull(Trim(!Mix2)), "", Trim(!Mix2))
                    
                End If
                .filter = ""
                .filter = ""
                .MoveFirst
                t = 0
                For i = 1 To .RecordCount
                
                    
                    If (InStr(!Recipe, Recipe) And Recipe <> "") Or ((InStr(!Mix1, Mix1) And Mix1 <> "") And (InStr(!Mix2, Mix2) And Mix2 <> "")) Then
                    
                        
                        ReDim Preserve uHannaCode(t)
                    
                        uHannaCode(t).Code = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
                        uHannaCode(t).Line = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
                        uHannaCode(t).STD = IIf(IsNull(Trim(!STD)), "", Trim(!STD))
                        uHannaCode(t).LoadInPrint = IIf(IsNull(Trim(!LoadInPrint)), "", Trim(!LoadInPrint))
                        uHannaCode(t).ProductName = IIf(IsNull(Trim(!ProductName)), "", Trim(!ProductName))
                        uHannaCode(t).Recipe = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
                        uHannaCode(t).Mix1 = IIf(IsNull(Trim(!Mix1)), "", Trim(!Mix1))
                        uHannaCode(t).Mix2 = IIf(IsNull(Trim(!Mix2)), "", Trim(!Mix2))
                        uHannaCode(t).Um = IIf(IsNull(Trim(!Um)), "", Trim(!Um))
                        uHannaCode(t).Qty = CheckDot(IIf(IsNull(Trim(!Qty)), "", Trim(!Qty)))
                        uHannaCode(t).MinQty = CheckDot(IIf(IsNull(Trim(!MinQty)), "", Trim(!MinQty)))
                        uHannaCode(t).MaxQty = CheckDot(IIf(IsNull(Trim(!MaxQty)), "", Trim(!MaxQty)))
                        
                        t = t + 1
                    End If
cont:
                    .MoveNext
                Next
                
            End With
   
ERR_END:

    On Error GoTo 0
    Debug.Print UBound(uHannaCode)
   
    AddCodeInSTDPreparationGrid = rc
    Exit Function
ERR_ADD:
    rc = False
    MsgBox err.Description
    Resume Next

End Function


Public Function GetSTDPreparationDate(ByVal STDPreparationID As Long) As String
With dbTabSTDPreparation
    .filter = ""
    .filter = "ID='" & STDPreparationID & "'"
    If .EOF Then
    Else
        GetSTDPreparationDate = IIf(IsNull(Trim(!startDate)), "", Trim(!startDate))
    End If


End With
End Function


Public Function IfNoPreparationRecipe(ByVal Recipe As String) As Boolean
Dim rc As Boolean

On Error GoTo ERR_IF

    rc = False
    
    With dbTabRecipe
        .filter = ""
        .filter = "Code='" & Trim(Recipe) & "'"
        If .EOF Then
        Else
            rc = !bNoPreparation
        End If
    
    End With

ERR_END:
    On Error GoTo 0
    IfNoPreparationRecipe = rc
    Exit Function
ERR_IF:
    rc = False
    MsgBox err.Description
    Resume Next

End Function

