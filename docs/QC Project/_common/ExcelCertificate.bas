Attribute VB_Name = "ExcelCertificate"
Option Explicit

' Aggiungi un riferimento a Microsoft Excel Object Library
' dal menu Progetto -> Riferimenti...

Private xlApp As Object
Private xlBook As Object
Private xlSheet As Object
Private chartObject As Excel.chartObject
Private a As Double
Private b As Double
Private x As Double
Private Y As Double
Public Function SetExcelCertificate_NEW(ByRef iCertificate As CertType, ByVal StettingName As String, ByRef ExcelFileName As String, ByVal FileSavedName As String) As Boolean
    ' Imposta i valori di a e b
    
    Dim rc As Boolean
    Dim i As Integer
    Dim t As Integer
    Dim k As Integer
    Dim UserDecimal As String
    Dim file_name As String
    
    On Error GoTo ERR_SET:

    ' Crea un nuovo oggetto Excel
    
    
    
    Set xlApp = CreateObject("Excel.Application")
    rc = True
     
     
     file_name = App.Path & "\LotCertificateForWeb_NEW.xls"
     
    
     If VerifyFile(file_name) Then

        Set xlBook = xlApp.Workbooks.Open(file_name)

    Else
        MsgBox "File Error..."
        Exit Function
    End If
    
    Dim Rows As Integer
    Dim strA As String
    Dim strB As String
    Dim strC As String
    Dim strD As String
    
    Dim ValA As String
    Dim ValB As String
    Dim ValC As String
    
    CloseSettingDataFile

   
     '"Certificate", "LplimGrph" , "UplimGrph"
   
     Set xlSheet = xlBook.Worksheets(3)
    
  
    Rows = GetSettingData(SettingName, "Certificate", "GrphCount", 0)
     
    If Rows > 0 Then
        For k = 1 To Rows
            strA = "D" & 1 + k
            strB = "E" & 1 + k
            strC = "A" & 1 + k
            ValA = GetSettingData(SettingName, "Certificate", "LplimGrph" & k, "")
            ValB = GetSettingData(SettingName, "Certificate", "UplimGrph" & k, "")
            ValC = GetSettingData(SettingName, "Certificate", "TargetValue" & k, "")
            
            xlSheet.Range(strA).Value = Replace(ValA, ",", ".")
            xlSheet.Range(strB).Value = Replace(ValB, ",", ".")
            xlSheet.Range(strC).Value = Replace(ValC, ",", ".")
            
            
        
        Next
        
    End If
        
    
    
    
    UserDecimal = GetSettingData(SettingName, "Certificate", "UserDecimal", "#0.000")
    
    
    ' Seleziona il foglio di lavoro "Certificate"
    Set xlSheet = xlBook.Worksheets(1)

    
    'Certificate Description
    
    With iCertificate
    
        xlSheet.Range("D4").Value = .ProductName
        xlSheet.Range("D5").Value = .ProductCode
        xlSheet.Range("D6").Value = .Method
        
        xlSheet.Rows("6:6").EntireRow.AutoFit
        
        xlSheet.Range("D7").Value = .RangePPM
        xlSheet.Range("D8").Value = .LotNumber
        xlSheet.Range("D9").Value = .BestUseBefore
        xlSheet.Range("D10").Value = .DateAnalisys
        xlSheet.Range("D11").Value = .ReferenceMeter
        xlSheet.Range("D12").Value = .ReferenceSTD
        xlSheet.Range("D13").Value = .Wavelenght
        xlSheet.Range("D14").Value = .CellMM
        
        xlSheet.Range("D34").Value = .RefSTDNote1 & " " & .RefSTDNote2
        xlSheet.Range("D36").Value = .ReferenceMeterDescription
        
        
        
        ' add Range Formula!
        
        
        xlSheet.Range("B7").Value = "Range [ " & .RangeFormula & " ]"
        xlSheet.Range("B21").Value = "Lot Result [ " & .RangeFormula & " ]"
        xlSheet.Range("F25").Value = "Standard Deviation [ " & .RangeFormula & " ]"
        xlSheet.Range("F27").Value = "Confidence interval (95%)[ " & .RangeFormula & " ]"
    
        FormatChemicalFormula xlSheet.Range("B7")
        FormatChemicalFormula xlSheet.Range("B21")
        FormatChemicalFormula xlSheet.Range("F25")
        FormatChemicalFormula xlSheet.Range("F27")
        
    End With
    
    
    'Certificate - Lot Result
    
   Rows = GetSettingData(SettingName, "Certificate - Lot Result", "Rows", 0)
   If Rows > 0 Then
   
       For k = 1 To Rows
            strA = "B" & 22 + k
            strB = "C" & 22 + k
            
            
            ValA = GetSettingData(SettingName, "Certificate - Lot Result", "StdValue" & k, "")
            ValB = GetSettingData(SettingName, "Certificate - Lot Result", "AverageResult" & k, "")
            
            ValA = Format(ValA, UserDecimal)
            ValB = Format(ValB, UserDecimal)
            xlSheet.Range(strA).Value = Replace(ValA, ",", ".")
            xlSheet.Range(strB).Value = Replace(ValB, ",", ".")
       Next
   End If
   
   
   'COMPONENTS IDENTIFICATION
    Dim iStart As Integer
    Dim rng As Excel.Range
    
    iStart = 67 'C
    
    For i = 1 To 5
        Select Case i
            Case 1
                strA = "D17"
                strB = "D18"
                strC = "D19"
                strD = "F19"
            Case 2
                strA = "G17"
                strB = "G18"
                strC = "G19"
                strD = "G19"
            Case 3
                strA = "H17"
                strB = "H18"
                strC = "H19"
                strD = "I19"
            Case 4
                strA = "J17"
                strB = "J18"
                strC = "J19"
                strD = "K19"
            Case 5
                strA = "L17"
                strB = "L18"
                strC = "L19"
                strD = "M19"
        End Select
          
            
            
            ValA = GetSettingData(SettingName, "Certificate - Components identification", "Code #" & i, "")
            ValB = GetSettingData(SettingName, "Certificate - Components identification", "Lot #" & i, "")
            ValC = GetSettingData(SettingName, "Certificate - Components identification", "Exp #" & i, "")
            
            If ValA = "" And ValB = "" And ValC = "" Then
            
            Else
            

                ' Supponiamo che strA, strB e strC siano celle adiacenti
                Set rng = xlSheet.Range(strA & ":" & strD)

                xlSheet.Range(strA).Value = Replace(ValA, ",", ".")
                xlSheet.Range(strB).Value = Replace(ValB, ",", ".")
                xlSheet.Range(strC).Value = Replace(ValC, ",", ".")
                
                ' Aggiungi bordi intorno al range
                With rng.Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With

            End If
    Next
   
   
   
    
   
   
   
    'Certificate - Calibration Function
    
    Rows = GetSettingData(SettingName, "Certificate - Calibration Function", "Rows", 0)
       ' SaveSettingData SettingName, "Certificate - Calibration Function", "Cols", .Cols - 1
        
    For k = 1 To 4
        
        If Rows > 0 Then
            strA = "L" & 23 + k
            strB = "M" & 23 + k
           ' strC = "L" & 16 + k
            
            
            ValA = GetSettingData(SettingName, "Certificate - Calibration Function", "Cell(" & k & ",1)", "")
            ValB = GetSettingData(SettingName, "Certificate - Calibration Function", "Cell(" & k & ",2)", "")
           ' ValC = GetSettingData(SettingName, "Certificate - Calibration Function", "Cell(" & k & ",3)", "")
            
            If k = 1 Then ' è un numero intero
                xlSheet.Range(strA).Value = Replace(ValA, ",", ".")
                xlSheet.Range(strB).Value = Replace(ValB, ",", ".")
            Else
                xlSheet.Range(strA).Value = Format(Replace(ValA, ",", "."), "0.000")
                xlSheet.Range(strB).Value = Format(Replace(ValB, ",", "."), "0.000")
            End If

            
         End If
    Next
    
     
    ValA = GetSettingData(SettingName, "Certificate - Calibration Function", "Cell(7,2)", "")
    xlSheet.Range("K21").Value = Format(Replace(ValA, ",", "."), "0.000")

   
   ' File number:
    xlSheet.Range("D40").Value = "CERT_ & SettingName"


    CloseSettingDataFile


    ' Stampa il foglio di lavoro "Certificate"
    'xlSheet.PrintOut Copies:=1, Collate:=True
    
    Dim MyExcelName As String
    Dim MyPfdName As String

    MyExcelName = USER_EXCEL_PATH & "\" & FormatNomeFile("CERT_" & FileSavedName & ".xls")
    MyPfdName = USER_EXCEL_PATH & "\" & FormatNomeFile("CERT_" & FileSavedName & ".pdf")
    
    



        Call SetWindowsPDFPrinter("")

        Dim w As New WshNetwork
        w.SetDefaultPrinter (PDFPrinterName)
        Set w = Nothing

    


    
    
    ' Salva il foglio di lavoro con un nuovo nome
    xlBook.SaveAs MyExcelName

    ' Esporta il foglio di lavoro come PDF
     xlSheet.ExportAsFixedFormat type:=xlTypePDF, FileName:=MyPfdName

    ' Chiudi il foglio di lavoro senza salvare le modifiche
    xlBook.Close SaveChanges:=False

    ' Chiudi l'applicazione Excel
    xlApp.Quit

    ' Rilascia gli oggetti
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
ERR_END:
    On Error GoTo 0
    ExcelFileName = MyExcelName
    SetExcelCertificate_NEW = rc
    Exit Function
    
ERR_SET:
    MsgBox Err.Description
    rc = False
    Resume Next
End Function


Public Function SetExcelCertificate(ByRef iCertificate As CertType, ByVal StettingName As String, ByRef ExcelFileName As String) As Boolean
    ' Imposta i valori di a e b
    
    Dim rc As Boolean
    Dim i As Integer
    Dim t As Integer
    Dim k As Integer
    
    On Error GoTo ERR_SET:

    ' Crea un nuovo oggetto Excel
   ' Set xlApp = New Excel.Application
    Set xlApp = CreateObject("Excel.Application")
    
    ' Apri il foglio di lavoro esistente
    
    
   'MsgBox App.PATH & "\LotCertificateForWeb.xls"
   
   
    rc = True
  
  MsgBox "1"
    
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\LotCertificateForWeb.xls")

  MsgBox "2"
    
    Dim Rows As Integer
    Dim strA As String
    Dim strB As String
    Dim strC As String
    
    Dim ValA As String
    Dim ValB As String
    Dim ValC As String
    
    CloseSettingDataFile

   
     '"Certificate", "LplimGrph" , "UplimGrph"
  
     Set xlSheet = xlBook.Worksheets("Lot Calculation")
     
     
  MsgBox "3"
  
    Rows = GetSettingData(SettingName, "Certificate", "GrphCount", 0)
     
    If Rows > 0 Then
        For k = 1 To Rows
            strA = "D" & 1 + k
            strB = "E" & 1 + k
            strC = "A" & 1 + k
            ValA = GetSettingData(SettingName, "Certificate", "LplimGrph" & k, "")
            ValB = GetSettingData(SettingName, "Certificate", "UplimGrph" & k, "")
            ValC = GetSettingData(SettingName, "Certificate", "TargetValue" & k, "")
            
            xlSheet.Range(strA).Value = Replace(ValA, ",", ".")
            xlSheet.Range(strB).Value = Replace(ValB, ",", ".")
            xlSheet.Range(strC).Value = Replace(ValC, ",", ".")
            
            
        
        Next
        
    End If
        
        
    ' Seleziona il foglio di lavoro "Certificate"
    Set xlSheet = xlBook.Worksheets("Certificate")

    
    'Certificate Description
    
    With iCertificate
    
        xlSheet.Range("F4").Value = .ProductName
        xlSheet.Range("F5").Value = .ProductCode
        xlSheet.Range("F6").Value = .Method
        
        xlSheet.Range("F7").Value = .RangePPM
        xlSheet.Range("F8").Value = .LotNumber
        xlSheet.Range("F9").Value = .BestUseBefore
        xlSheet.Range("F10").Value = .DateAnalisys
        xlSheet.Range("F11").Value = .ReferenceMeter
        xlSheet.Range("F12").Value = .ReferenceSTD
        xlSheet.Range("F13").Value = .Wavelenght
        xlSheet.Range("F14").Value = .CellMM
        
        xlSheet.Range("D34").Value = .RefSTDNote1
    
    End With
    
    
    'Certificate - Lot Result
    
   Rows = GetSettingData(SettingName, "Certificate - Lot Result", "Rows", 0)
   If Rows > 0 Then
   
       For k = 1 To Rows
            strA = "E" & 17 + k
            strB = "F" & 17 + k
            
            
            ValA = GetSettingData(SettingName, "Certificate - Lot Result", "StdValue" & k, "")
            ValB = GetSettingData(SettingName, "Certificate - Lot Result", "AverageResult" & k, "")
            
            xlSheet.Range(strA).Value = Replace(ValA, ",", ".")
            xlSheet.Range(strB).Value = Replace(ValB, ",", ".")
       Next
   End If
   
   
    'Certificate - Calibration Function
    
    Rows = GetSettingData(SettingName, "Certificate - Calibration Function", "Rows", 0)
       ' SaveSettingData SettingName, "Certificate - Calibration Function", "Cols", .Cols - 1
        
    For k = 1 To Rows
        
        If Rows > 0 Then
            strA = "J" & 16 + k
            strB = "K" & 16 + k
            strC = "L" & 16 + k
            
            
            ValA = GetSettingData(SettingName, "Certificate - Calibration Function", "Cell(" & k & ",1)", "")
            ValB = GetSettingData(SettingName, "Certificate - Calibration Function", "Cell(" & k & ",2)", "")
            ValC = GetSettingData(SettingName, "Certificate - Calibration Function", "Cell(" & k & ",3)", "")
            
            
            xlSheet.Range(strA).Value = Replace(ValA, ",", ".")
            xlSheet.Range(strB).Value = Replace(ValB, ",", ".")
            xlSheet.Range(strC).Value = Replace(ValC, ",", ".")
            
         End If
    Next
     
  
            


    CloseSettingDataFile


    ' Stampa il foglio di lavoro "Certificate"
    'xlSheet.PrintOut Copies:=1, Collate:=True
    
    Dim MyExcelName As String
    Dim MyPfdName As String

    MyExcelName = USER_EXCEL_PATH & "\" & FormatNomeFile("CERT_" & SettingName & ".xls")
    MyPfdName = USER_EXCEL_PATH & "\" & FormatNomeFile("CERT_" & SettingName & ".pdf")
    
    



        Call SetWindowsPDFPrinter("")

        Dim w As New WshNetwork
        w.SetDefaultPrinter (PDFPrinterName)
        Set w = Nothing

    


    
    
    ' Salva il foglio di lavoro con un nuovo nome
    xlBook.SaveAs MyExcelName

    ' Esporta il foglio di lavoro come PDF
    xlSheet.ExportAsFixedFormat type:=xlTypePDF, FileName:=MyPfdName

    ' Chiudi il foglio di lavoro senza salvare le modifiche
    xlBook.Close SaveChanges:=False

    ' Chiudi l'applicazione Excel
    xlApp.Quit

    ' Rilascia gli oggetti
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
ERR_END:
    On Error GoTo 0
    ExcelFileName = MyExcelName
    SetExcelCertificate = rc
    Exit Function
    
ERR_SET:
    MsgBox Err.Description
    rc = False
    Resume Next
    
    
        ' Genera un grafico pivot della funzione y=a+bx
    ' Questo è un esempio e potrebbe non funzionare per il tuo caso specifico
    ' Adattalo alle tue esigenze
    'For X = 1 To 10
    '    Y = a + b * X
    '    xlSheet.Cells(X, 1).Value = X
    '    xlSheet.Cells(X, 2).Value = Y
    'Next X
    ' Crea un nuovo grafico
    'Set chartObject = xlSheet.ChartObjects.Add(Left:=100, Width:=375, Top:=50, Height:=225)
    'chartObject.Chart.SetSourceData Source:=xlSheet.Range("A1:B10")
    'chartObject.Chart.ChartType = xlLine
End Function


Private Sub FormatChemicalFormula_old(cell As Range)
    Dim i As Integer
    Dim text As String
    Dim inBrackets As Boolean
    text = cell.Value
    inBrackets = False

    For i = 1 To Len(text)
        If Mid(text, i, 1) = "[" Then
            inBrackets = True
        ElseIf Mid(text, i, 1) = "]" Then
            inBrackets = False
        End If

        If inBrackets And (IsNumeric(Mid(text, i, 1)) Or Mid(text, i, 1) = "+" Or Mid(text, i, 1) = "-") Then
            cell.Characters(start:=i, Length:=1).Font.Superscript = True
        End If
    Next i
End Sub

' Esempio di utilizzo
'xlSheet.Range("D6").Value = "Confidence interval (95%)[NH4+]"
'FormatChemicalFormula xlSheet.Range("D6")


Private Sub FormatChemicalFormula(cell As Range)
    Dim i As Integer
    Dim text As String
    Dim inBrackets As Boolean
    text = cell.Value
    inBrackets = False

    For i = Len(text) To 1 Step -1
        If Mid(text, i, 1) = "]" Then
            inBrackets = True
        ElseIf Mid(text, i, 1) = "[" Then
            inBrackets = False
        End If

        If inBrackets Then
            If IsNumeric(Mid(text, i, 1)) Then
                If i < Len(text) And (Mid(text, i + 1, 1) = "+" Or Mid(text, i + 1, 1) = "-") Then
                    cell.Characters(start:=i, Length:=1).Font.Superscript = True
                Else
                    cell.Characters(start:=i, Length:=1).Font.Subscript = True
                End If
            ElseIf Mid(text, i, 1) = "+" Or Mid(text, i, 1) = "-" Then
                cell.Characters(start:=i, Length:=1).Font.Superscript = True
            End If
        End If
    Next i
End Sub

' Esempio di utilizzo
'xlSheet.Range("D6").Value = "Confidence interval (95%)[NH4+]"
'FormatChemicalFormula xlSheet.Range("D6")

