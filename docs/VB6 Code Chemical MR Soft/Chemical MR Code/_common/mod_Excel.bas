Attribute VB_Name = "mod_Excel"
Option Explicit
'Public ExcelApp As Excel.Application
'Public Cartella As Excel.Workbook
'Public Foglio As Excel.Worksheet

Public ExcelApp As Object

Public Cartella As Object
Public Foglio As Object

Private SettingName As String

Public Function dbChemicalMRToExcel(ByRef pBar As ProgressBar) As Boolean
Dim rc As Boolean
Dim sString As String
On Error GoTo ERR_EXP
    rc = True
    pBar.Value = 0
    pBar.Visible = True
    sString = FormatNomeFile("MR data base " & FormatDateTime(Now, vbShortDate) & " rev 2")
    If CreateExcel(False) Then
        NewExcelWorksheet (sString)
        If CopyHannaCodeData(pBar) Then
            Call SaveHannaCodeExcel(sString)
            Call CloseExcel
            PopupMessage 2, "Excel file correctly generated..." & vbCrLf & USER_DESKTOP & "\" & FormatNomeFile(sString) & ".xlsx"
        Else
            rc = False
        End If
    Else
        rc = False
    End If
        
        
    
ERR_END:
    dbChemicalMRToExcel = rc
    pBar.Visible = False
    Exit Function
ERR_EXP:
    rc = False
    MsgBox Err.Description
    Resume ERR_END:
End Function

Private Function SaveHannaCodeExcel(ByVal sString As String)
Dim Mystring As String
    Mystring = USER_DESKTOP & "\" & FormatNomeFile(sString) & ".xlsx"
    ExcelApp.DisplayAlerts = False
    ExcelApp.ActiveWorkbook.SaveAs FileName:=Mystring
End Function
Public Function DBCodeToExcel(ByRef pBar As ProgressBar, ByVal Index As Integer, ByVal DatabaseName As String) As Boolean
Dim rc As Boolean
Dim sString As String
On Error GoTo ERR_EXP
    rc = True
    pBar.Value = 0
    pBar.Visible = True
    sString = FormatNomeFile(Replace("MR XXX data base ", "XXX", DatabaseName) & FormatDateTime(Now, vbShortDate) & " rel" & dbCodeRelease)
    If CreateExcel(False) Then
            NewExcelWorksheet (DatabaseName)
            Select Case Index
                Case 0
                    ' hanna code
                    rc = True
                    Call ExportAccessToExcel(sString, pBar)
                    PopupMessage 2, "Excel file correctly generated..." & vbCrLf & USER_DESKTOP & "\" & FormatNomeFile(sString) & ".xlsx"
                    Exit Function
                Case 1
                    
                Case 2
     
                Case 3
        
                Case 4
     
              
                Case 5
                    ' chemical MR
                    rc = CopyChemicalMRData(pBar)
            
            End Select
        
    Else
        rc = False
    End If
    
        
    If rc Then
            
            Call SaveExcelFileOnDesktop(sString)
            Call CloseExcel
            PopupMessage 2, "Excel file correctly generated..." & vbCrLf & USER_DESKTOP & "\" & FormatNomeFile(sString) & ".xlsx"
               
    
    End If
    
    
        
    
ERR_END:
    DBCodeToExcel = rc
    pBar.Visible = False
    Exit Function
ERR_EXP:
    rc = False
    MsgBox Err.Description
    Resume ERR_END:
End Function


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



Public Function SaveExcel(ByVal sString As String, Optional PATH As String)
Dim Mystring As String

    If PATH = "" Then PATH = USER_EXCEL_PATH
    
    If Len(sString) > 27 Then sString = Left$(sString, 27)
   
    Mystring = PATH & "\" & FormatNomeFile(sString) & ".xlsx"


    ExcelApp.DisplayAlerts = False
    'ExcelApp.Visible = True
   
    ExcelApp.ActiveWorkbook.SaveAs FileName:=Mystring 'MyWeightCheck.Code & MyWeightCheck.Lot '(USER_EXCEL_PATH & "\" & (MyWeightCheck.Code & MyWeightCheck.Lot))
    'ExcelApp.SaveWorkspace 'USER_EXCEL_PATH & "\" & MyWeightCheck.FileName
End Function


Public Function SaveExcelFileOnDesktop(ByVal sString As String)
Dim Mystring As String
    Mystring = USER_DESKTOP & "\" & FormatNomeFile(sString) & ".xlsx"
    ExcelApp.DisplayAlerts = False
    ExcelApp.ActiveWorkbook.SaveAs FileName:=Mystring
End Function
Public Function CreateExcel(Optional ByVal bValue As Boolean = True) As Boolean
    Dim rc As Boolean
    On Error GoTo ERR_CREATE
    rc = True
   ' Set ExcelApp = New Excel.Application
     Set ExcelApp = CreateObject("Excel.Application")

    
    'ExcelApp.Visible = True 'bvalue
ERR_END:
    On Error GoTo 0
    CreateExcel = rc
    Exit Function
ERR_CREATE:
    rc = False
    MsgBox Err.Description
    Resume ERR_END
End Function
    
Public Sub NewExcelWorksheet(sString As String)
    
    On Error GoTo ERR_EXC
    Set Cartella = ExcelApp.Workbooks.Add
    Set Foglio = Cartella.Worksheets(1)
    
    Foglio.Name = "Preparation Details"
    
ERR_END:
    Exit Sub
ERR_EXC:
    MsgBox Err.Description
    
End Sub
Public Function AddCodeValue(Col As Long, Row As Integer, Text As String) As Boolean
    Dim rc As Boolean
    Dim bValue As Boolean
    Dim Intervallo As Excel.Range
    Dim sRange As String


    On Error GoTo ERR_ADD
    On Error GoTo ERR_ADD
    rc = True

    Set Foglio = Cartella.Worksheets(1)
    
    Select Case Row
        Case Is <= 26
             sRange = Chr(64 + Row) & Col
        Case 27 To 52
            sRange = "A" & Chr(64 + Row - 26) & Col
        Case 53 To 78
            sRange = "B" & Chr(64 + Row - 52) & Col
        Case 79 To 104
            sRange = "C" & Chr(64 + Row - 78) & Col
    
    End Select
   
    'If Row > 26 Then
    '    If Row > 52 Then
     '       sRange = "B" & Chr(64 + Row - 52) & Col
     '   Else
     '
     '       sRange = "A" & Chr(64 + Row - 26) & Col
     '   End If
        
     '   Debug.Print sRange
        
    'Else
     '   sRange = Chr(64 + Row) & Col
    'End If
    
    
    
    Set Intervallo = Foglio.Range(sRange)
    
    Intervallo.Value = Text

ERR_END:
    On Error GoTo 0
    AddCodeValue = rc
    Exit Function
ERR_ADD:

    Resume Next
End Function
Public Function AddValue(Col As Integer, Row As Integer, Text As String, Optional USER_COLOR As Boolean = False, Optional bTitle As Boolean = False, Optional Color As OLE_COLOR, Optional bNonSelezionati As Boolean, Optional bWrap As Boolean) As Boolean
    Dim rc As Boolean
    Dim bValue As Boolean
    Dim Intervallo As Excel.Range
    Dim sRange As String
    Dim COLOR_BACK As OLE_COLOR
    Dim COLOR_FORE As OLE_COLOR
    Dim COLOR_BACK_NONSELEZIONATI As OLE_COLOR


    On Error GoTo ERR_ADD
    rc = True

    bValue = False
    
    COLOR_BACK = &H886010   'vbColorTextLightBlue
    COLOR_FORE = vbWhite
    COLOR_BACK_NONSELEZIONATI = &H8080FF
    Set Foglio = Cartella.Worksheets(1)
   
    If Row > 26 Then
        sRange = "A" & Chr(64 + Row - 26) & Col
    Else
        sRange = Chr(64 + Row) & Col
    End If
    Set Intervallo = Foglio.Range(sRange)
    
    If bWrap Then Intervallo.WrapText = True

   ' If IIntervallosDate(CDate(Text)) Then
   '     If Len(Trim(Text)) = 10 Then
   '         Intervallo.NumberFormat = "yyyy/mm/dd"
   '     ElseIf Len(Trim(Text)) = 7 Then
   '         Intervallo.NumberFormat = "mm/yyyy"
   '     ElseIf InStr(Text, ":") Then
   '         Text = "'" & Text
   '     End If
   ' End If
    
    

    Intervallo.Value = (Text)

    Call BoxIt(Intervallo)

 
 
    If bTitle Then
        Intervallo.Font.Size = 14
        Intervallo.Font.Bold = True
    End If
    
    If bNonSelezionati Then
        
        Intervallo.Interior.Color = COLOR_BACK_NONSELEZIONATI
    End If
    
    If USER_COLOR Then
        bValue = True
        Intervallo.Interior.Color = COLOR_BACK ' RGB(200, 160, 35)
        Intervallo.Font.Color = COLOR_FORE  ' RGB(200, 160, 35)
        
    Else
        
    End If
    If Color <> 0 Then
        Intervallo.Font.Color = Color
    End If
    
    Intervallo.Font.Bold = bValue
   
ERR_END:
    On Error GoTo 0
    AddValue = rc
    Exit Function
ERR_ADD:
   ' rc = False
    'MsgBox err.Description
    Resume Next
End Function
Private Sub BoxIt(aRng As Range)
On Error Resume Next

    With aRng

        'Clear existing
        .Borders.LineStyle = xlNone

        'Apply new borders
        .BorderAround xlContinuous, xlThin, 0
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .Weight = xlThin
        End With
    End With

End Sub
Public Function CloseExcel() As Boolean
    Dim rc As Boolean
    On Error GoTo ERR_CLOSE
    rc = True
    ExcelApp.Quit
    Set ExcelApp = Nothing
ERR_END:
    On Error GoTo 0
    CloseExcel = rc
    Exit Function
ERR_CLOSE:
    rc = False
    MsgBox Err.Description
    Resume Next
End Function
Public Function FormatPage()
    Dim rc As Boolean
    Dim Intervallo As Excel.Range
    Dim sRange As String
    
    On Error GoTo ERR_ADD
    rc = True
    Set Foglio = Cartella.Worksheets(1)
    
    With Foglio
    
        .Rows("1:1").Select

        
        With .Range("A:CZ")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            
           ' .NumberFormat = "#,##0.0000"
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = True
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
              
              
     


        With .Range("A1", "Z1000")
            .RowHeight = 18
        End With
        
        .Columns("B:CZ").ColumnWidth = 20

     
    End With

    
ERR_END:
    On Error GoTo 0
    FormatPage = rc
    Exit Function
ERR_ADD:
    rc = False
    MsgBox Err.Description
    Resume ERR_END
End Function
