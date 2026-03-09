Attribute VB_Name = "mod_Excel_Esporta_Database"
Option Explicit


Public Function CopyHannaCodeData(ByRef pBar As ProgressBar) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim strDecimal As String

    On Error GoTo ERR_COPY
    
    FormatPage
    
    rc = True
    i = 0
    
    With dbTabCode
        .filter = ""
        If .EOF Then
            rc = False
            GoTo ERR_END:
        Else
            .MoveFirst
            pBar.Max = .RecordCount
            
            For t = 1 To .fields.Count - 1
                   Call AddCodeValue(1, t + 2, IIf(IsNull(Trim(.fields(t).Name)), "", Trim(.fields(t).Name)))
            Next
                
            Do
                i = i + 1
                For t = 1 To .fields.Count - 1
                    If t = 17 Or t = 18 Then
                        If InStr(Trim(.fields(t)), "/") Then
                            GoTo TrueValue
                        Else
                        Call AddCodeValue(i + 1, t + 2, IIf(IsNull(Trim(.fields(t))), "", Trim(.fields(t)) & "%"))
                        End If
                    ElseIf t = 50 Then
                        If InStr(Trim(UCase(.fields(t))), "FALSE") Then
                        Else
                            Call AddCodeValue(i + 1, t + 2, IIf(IsNull(Trim(.fields(t))), "", Trim(.fields(t))))
                        End If
                    Else
TrueValue:
                        Call AddCodeValue(i + 1, t + 2, IIf(IsNull(Trim(.fields(t))), "", Trim(.fields(t))))
                    End If
                Next
                .MoveNext
                pBar.Value = i
            Loop Until .EOF
        End If
    
    
    End With
ERR_END:
    CopyHannaCodeData = rc
    Exit Function
ERR_COPY:
    rc = False
    MsgBox err.Description
    GoTo ERR_END:
End Function


Private Function CopyChemicalMRData(ByRef pBar As ProgressBar) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim strDecimal As String

    On Error GoTo ERR_COPY
    
    FormatPage
    
    rc = True
    i = 0
    
    With dbTabRawMaterial
        .filter = ""
        If .EOF Then
            rc = False
            GoTo ERR_END:
        Else
            .MoveFirst
            pBar.Max = .RecordCount
            
            For t = 1 To .fields.Count - 1
                   Call AddCodeValue(1, t + 2, IIf(IsNull(Trim(.fields(t).Name)), "", Trim(.fields(t).Name)))
            Next
                
            Do
                i = i + 1
                For t = 1 To .fields.Count - 1
TrueValue:
                        Call AddCodeValue(i + 1, t + 2, IIf(IsNull(Trim(.fields(t))), "", Trim(.fields(t))))
                Next
                .MoveNext
                pBar.Value = i
            Loop Until .EOF
        End If
    
    
    End With
ERR_END:
    CopyChemicalMRData = rc
    Exit Function
ERR_COPY:
    rc = False
    MsgBox err.Description
    GoTo ERR_END:
End Function

Public Function CopySTDPreparationWayData(ByRef pBar As ProgressBar) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim strDecimal As String

    On Error GoTo ERR_COPY
    
    FormatPage
    
    rc = True
    i = 0
    
    With dbTabSTDPreparationWay
        .filter = ""
        If .EOF Then
            rc = False
            GoTo ERR_END:
        Else
            .MoveFirst
            pBar.Max = .RecordCount
            
            For t = 1 To .fields.Count - 1
                   Call AddCodeValue(1, t + 2, IIf(IsNull(Trim(.fields(t).Name)), "", Trim(.fields(t).Name)))
            Next
                
            Do
                i = i + 1
                For t = 1 To .fields.Count - 1
TrueValue:
                        Call AddCodeValue(i + 1, t + 2, IIf(IsNull(Trim(.fields(t))), "", Trim(.fields(t))))
                Next
                .MoveNext
                pBar.Value = i
            Loop Until .EOF
        End If
    
    
    End With
ERR_END:
    CopySTDPreparationWayData = rc
    Exit Function
ERR_COPY:
    rc = False
    MsgBox err.Description
    GoTo ERR_END:
End Function


Public Function CopyCodeClassificationData(ByRef pBar As ProgressBar) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim strDecimal As String

    On Error GoTo ERR_COPY
    
    FormatPage
    
    rc = True
    i = 0
    
    With dbTabCodeClassification
        .filter = ""
        If .EOF Then
            rc = False
            GoTo ERR_END:
        Else
            .MoveFirst
            pBar.Max = .RecordCount
            
            For t = 1 To .fields.Count - 1
                   Call AddCodeValue(1, t + 2, IIf(IsNull(Trim(.fields(t).Name)), "", Trim(.fields(t).Name)))
            Next
                
            Do
                i = i + 1
                For t = 1 To .fields.Count - 1
TrueValue:
                        Call AddCodeValue(i + 1, t + 2, IIf(IsNull(Trim(.fields(t))), "", Trim(.fields(t))))
                Next
                .MoveNext
                pBar.Value = i
            Loop Until .EOF
        End If
    
    
    End With
ERR_END:
    CopyCodeClassificationData = rc
    Exit Function
ERR_COPY:
    rc = False
    MsgBox err.Description
    GoTo ERR_END:
End Function



Public Function CopyFrasiHData(ByRef pBar As ProgressBar) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim strDecimal As String

    On Error GoTo ERR_COPY
    
    FormatPage
    
    rc = True
    i = 0
    
    With dbTabFrasiH
        .filter = ""
        If .EOF Then
            rc = False
            GoTo ERR_END:
        Else
            .MoveFirst
            pBar.Max = .RecordCount
            
            For t = 1 To .fields.Count - 1
                   Call AddCodeValue(1, t + 2, IIf(IsNull(Trim(.fields(t).Name)), "", Trim(.fields(t).Name)))
            Next
                
            Do
                i = i + 1
                For t = 1 To .fields.Count - 1
TrueValue:
                        Call AddCodeValue(i + 1, t + 2, IIf(IsNull(Trim(.fields(t))), "", Trim(.fields(t))))
                Next
                .MoveNext
                pBar.Value = i
            Loop Until .EOF
        End If
    
    
    End With
ERR_END:
    CopyFrasiHData = rc
    Exit Function
ERR_COPY:
    rc = False
    MsgBox err.Description
    GoTo ERR_END:
End Function




Public Function CopyRecipesData(ByRef pBar As ProgressBar) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim strDecimal As String

    On Error GoTo ERR_COPY
    
    FormatPage
    
    rc = True
    i = 0
    
    With dbTabRecipe
        .filter = ""
        If .EOF Then
            rc = False
            GoTo ERR_END:
        Else
            .MoveFirst
            pBar.Max = .RecordCount
            
            For t = 1 To .fields.Count - 1
                   Call AddCodeValue(1, t + 2, IIf(IsNull(Trim(.fields(t).Name)), "", Trim(.fields(t).Name)))
            Next
                
            Do
                i = i + 1
                For t = 1 To .fields.Count - 1
TrueValue:
                        Call AddCodeValue(i + 1, t + 2, IIf(IsNull(Trim(.fields(t))), "", Trim(.fields(t))))
                Next
                .MoveNext
                pBar.Value = i
            Loop Until .EOF
        End If
    
    
    End With
ERR_END:
    CopyRecipesData = rc
    Exit Function
ERR_COPY:
    rc = False
    MsgBox err.Description
    GoTo ERR_END:
End Function






