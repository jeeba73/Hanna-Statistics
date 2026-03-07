Attribute VB_Name = "mod_Excel_Esporta_Database"
Option Explicit

Public Sub ExportAccessToExcel(ByVal sString As String, ByRef pBar As ProgressBar)
    Dim objExcel As Object
    Dim objWorkbook As Object
    Dim objWorksheet As Object
    Dim objAccess As Object
    Dim rs As Object
    Dim dbPathAccess As String
    Dim tableName As String
    Dim excelPath As String
    Dim i As Integer
    Dim j As Integer
    
    Dim MaxCount As Long
    
    dbTabCode.filter = ""
    
    MaxCount = dbTabCode.RecordCount

    ' Percorso del database Access
    dbPathAccess = dbPath & dbCodeName
    ' Nome della tabella da esportare
    tableName = "TabCode"
    ' Percorso del file Excel di destinazione
    excelPath = USER_DESKTOP & "\" & FormatNomeFile(sString) & ".xlsx"  '"C:\Percorso\AlTuoFile.xlsx"

    ' Crea oggetti Access e Excel
    Set objAccess = CreateObject("DAO.DBEngine.36")
    
    Set objExcel = CreateObject("Excel.Application")

    ' Apri il database Access
    Set rs = objAccess.OpenDatabase(dbPathAccess).OpenRecordset("SELECT * FROM " & tableName)
    'Debug.Print dbPath

    ' Crea un nuovo workbook Excel
    Set objWorkbook = objExcel.Workbooks.Add
    Set objWorksheet = objWorkbook.Sheets(1)

    ' Copia i nomi delle colonne
    
    pBar.Min = 0
    pBar.Max = MaxCount + 100
    For i = 0 To rs.fields.Count - 1
        objWorksheet.Cells(1, i + 1).Value = rs.fields(i).Name
    Next i

    ' Copia i dati
    i = 2
    Do While Not rs.EOF
         pBar.Value = i
        For j = 0 To rs.fields.Count - 1
            objWorksheet.Cells(i, j + 1).Value = "'" & rs.fields(j).Value
        Next j
        rs.MoveNext
        i = i + 1
    Loop

    ' Salva il file Excel
    objWorkbook.SaveAs excelPath
    objWorkbook.Close
    objExcel.Quit

    ' Rilascia gli oggetti
    Set rs = Nothing
    Set objAccess = Nothing
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    
    pBar.Visible = False
  '  MsgBox "Esportazione completata!"
End Sub

Public Function CopyHannaCodeData(ByRef pBar As ProgressBar) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim X As Integer
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
            X = 1
            For t = 1 To .fields.Count - 1
                
               ' Select Case t
               '     Case 7, 8, 15, 16, 17, 18, 21, 22, 24, 25, 27, 28, 30, 31, 33, 34, 67
                '         GoTo cont:
               '    Case 36 To 51
               '          GoTo cont:
               ' End Select
                
                X = X + 1
                Call AddCodeValue(1, X + 2, IIf(IsNull(Trim(.fields(t).Name)), "", "'" & Trim(.fields(t).Name)))
cont:
            Next
                
            Do
                i = i + 1
                X = 1
                For t = 1 To .fields.Count - 1
                    
                      X = X + 1

TrueValue:
                        Call AddCodeValue(i + 1, X + 2, IIf(IsNull(Trim(.fields(t))), "", "'" & Trim(.fields(t))))
                    
cont2:
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


Public Function CopyChemicalRMData(ByRef pBar As ProgressBar) As Boolean
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
    CopyChemicalRMData = rc
    Exit Function
ERR_COPY:
    rc = False
    MsgBox err.Description
    GoTo ERR_END:
End Function

Public Function CopyProductionWayData(ByRef pBar As ProgressBar) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim strDecimal As String

    On Error GoTo ERR_COPY
    
    FormatPage
    
    rc = True
    i = 0
    
    With dbTabProductionWay
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
    CopyProductionWayData = rc
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






