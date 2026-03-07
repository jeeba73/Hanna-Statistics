Attribute VB_Name = "mod_Excel"
Option Explicit
'Public ExcelApp As Excel.Application
'Public Cartella As Excel.Workbook
'Public Foglio As Excel.Worksheet

Public ExcelApp As Object

Public Cartella As Object
Public Foglio As Object

Private SettingName As String


Public Function DBCodeToExcel(ByRef pBar As ProgressBar, ByVal Index As Integer, ByVal DatabaseName As String) As Boolean
Dim rc As Boolean
Dim sString As String
On Error GoTo ERR_EXP
    rc = True
    pBar.Value = 0
    pBar.Visible = True
    sString = FormatNomeFile(Replace("CP XXX data base ", "XXX", DatabaseName) & FormatDateTime(Now, vbShortDate) & " rel" & dbCodeRelease)
    If CreateExcel(False) Then
            NewExcelWorksheet (DatabaseName)
            Select Case Index
                Case 0
                    ' hanna code
                    rc = True
                    Call ExportAccessToExcel(sString, pBar)
                    PopupMessage 2, "Excel file correctly generated..." & vbCrLf & USER_DESKTOP & "\" & FormatNomeFile(sString) & ".xlsx"
                    Exit Function
                   ' rc = CopyHannaCodeData(pBar)
                Case 1
                    ' Production Way
                    rc = CopyProductionWayData(pBar)
                Case 2
                    ' Code Classification
                    rc = CopyCodeClassificationData(pBar)
                Case 3
                    ' Frasi H
                    rc = CopyFrasiHData(pBar)
                Case 4
                    ' Recipe
                    rc = CopyRecipesData(pBar)
              
                Case 5
                    ' Chemical RM
                    rc = CopyChemicalRMData(pBar)
            
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
    MsgBox err.Description
    Resume ERR_END:
End Function





Public Function SaveExcel(ByVal sString As String, Optional Path As String)
Dim Mystring As String

    If Path = "" Then Path = USER_EXCEL_PATH
    
     If Len(sString) > 40 Then sString = Left$(sString, 40)

    Mystring = Path & "\" & FormatNomeFile(sString) & ".xlsx"


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
    MsgBox err.Description
    Resume ERR_END
End Function
    
Public Sub NewExcelWorksheet(sString As String)
    
    On Error GoTo ERR_EXC
    Set Cartella = ExcelApp.Workbooks.Add
    Set Foglio = Cartella.Worksheets(1)
    
    If Len(sString) > 20 Then sString = Left$(sString, 20)
    
    Foglio.Name = FormatNomeFile(sString)
    
ERR_END:
    Exit Sub
ERR_EXC:
    MsgBox err.Description
    
End Sub
Public Function AddCodeValue(Col As Long, Row As Integer, Text As String) As Boolean
    Dim rc As Boolean
    Dim bValue As Boolean
    Dim Intervallo As Excel.Range
    Dim sRange As String


    On Error GoTo ERR_ADD
    rc = True

    Set Foglio = Cartella.Worksheets(1)
   
    If Row > 26 And Row < 51 Then
       ' Debug.Print Chr(64 + Row - 26)
        sRange = "A" & Chr(64 + Row - 26) & Col
    ElseIf Row > 51 And Row < 76 Then
       ' Debug.Print Chr(64 + Row - 51)
    
        sRange = "B" & Chr(64 + Row - 51) & Col
    ElseIf Row > 76 Then
       ' Debug.Print Chr(64 + Row - 77)
    
        sRange = "C" & Chr(64 + Row - 76) & Col
    Else
    
    
        sRange = Chr(64 + Row) & Col
    End If
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
    MsgBox err.Description
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

        
        With .Range("A:AZ")
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
        
        .Columns("B:AZ").ColumnWidth = 20

        .Columns("C:C").ColumnWidth = 40

    End With

    
ERR_END:
    On Error GoTo 0
    FormatPage = rc
    Exit Function
ERR_ADD:
    rc = False
    MsgBox err.Description
    Resume ERR_END
End Function
