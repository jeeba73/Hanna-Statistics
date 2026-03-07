Attribute VB_Name = "mod_Excel_importa_file"
Option Explicit







'variabile oggetto che contiene il riferimento alla cartella di lavoro di Excel
Private FileExcel As Object

'variabile oggetto che contiene il riferimento al foglio di lavoro di Excel
Private FoglioExcel As Object

'variabile oggetto che contiene il  riferimento alle celle del foglio di lavoro
'di Excel
Private CellaFoglioExcel As Range




Public Function HannaCodeExcelImport(ByVal mFile As String, ByVal Frm As Form, Optional ByRef CodeCount As Long, Optional ByVal bDeletePreviousRecords As Boolean) As Boolean
Dim rc As Boolean
Dim file_name As String


Dim i As Integer
Dim t As Integer
Dim x As Integer
Dim r As Long
Dim nMax As Integer
Dim strRecipe As String


Dim sDestinazione As String

On Error GoTo ERR_CREATE_OBJECT


If bDeletePreviousRecords Then
   
   Call DeleteAllTabCode

End If

On Error GoTo ERR_CREATE_OBJECT
    rc = True
    file_name = mFile

    If VerifyFile(file_name) Then

        'imposto la variabile oggetto FileExcel con il nome del file xls
        Set FileExcel = Excel.Workbooks.Open(file_name)
    Else
        Frm.List1.AddItem Now & " - Impossibile aprire il file."
        PopupMessage 2, "Impossibile aprire il file." & _
        vbCrLf & "Il file " & file_name & " non è stato trovato.", "Importa Excel", True
        rc = False
        GoTo END_FN
    End If
    Set FoglioExcel = FileExcel.Worksheets(1)
            
     
    Set FoglioExcel = FileExcel.Worksheets(1)
    
    
            
    r = 1
   
    i = 0
    Frm.List1.AddItem Now & " - Loading Hanna Code ..."
    

    Dim NewCode As Integer
    Dim sString As String
    Do
        i = i + 1
        r = r + 1

               
        MyImportHannaCode.code = Trim(FoglioExcel.Cells(r, 2))
        MyImportHannaCode.ProductName = Trim(FoglioExcel.Cells(r, 5))
        MyImportHannaCode.RangeMin = Trim(FoglioExcel.Cells(r, 30))
        MyImportHannaCode.RangeMax = Trim(FoglioExcel.Cells(r, 31))
        
      
        
        If MyImportHannaCode.code = "" Then
            GoTo END_FN
        End If
        
        With dbTabCode
            .Filter = ""
            
            If MyImportHannaCode.RangeMin = "" Or MyImportHannaCode.RangeMax = "" Then
                 .Filter = "Code='" & MyImportHannaCode.code & "'"
            Else
                .Filter = "Code='" & MyImportHannaCode.code & "' and RangeMin='" & MyImportHannaCode.RangeMin & "' and RangeMax='" & MyImportHannaCode.RangeMax & "'"
            End If
            If Not (.EOF) Then
                
                ' ok controllo che sia associata al cliente
                Frm.List1.AddItem Now & " - Hanna SFG Code (" & i & ") : " & MyImportHannaCode.code & " ( " & MyImportHannaCode.ProductName & " ) already Exsists... "
            Else
                .AddNew
                NewCode = NewCode + 1
                Frm.List1.AddItem Now & " - Import new Hanna SFG Code (" & i & ") : " & MyImportHannaCode.code & " ( " & MyImportHannaCode.ProductName & " )"
            End If
                
            For t = 1 To .Fields.Count - 1

                   If InStr(.Fields(t).Name, "Date") Then
                     !DateModified = Now
                 
                   Else
                    'If .Fields(t) = Trim(FoglioExcel.Cells(r, t + 1)) Then
                   ' Else
                        'Debug.Print .Fields(t).Value
                    'End If
                    .Fields(t).Value = Trim(FoglioExcel.Cells(r, t + 1))
                End If
       
            Next
            
            
               
                Frm.List1.AddItem Now & " - Hanna SFG Code (" & i & ") : " & MyImportHannaCode.code & " ( " & MyImportHannaCode.ProductName & " ) Saved... "
                DoEvents
                .Update
                
              
        End With
    Loop Until (FoglioExcel.Cells(r + 1, 2) = "" And FoglioExcel.Cells(r + 2, 2) = "")
    
    CodeCount = i
    Frm.List1.AddItem ""
    Frm.List1.AddItem "n." & NewCode & " New Hanna Code Imported"
    Frm.List1.AddItem "n." & CodeCount & " Excel Code"
    Frm.List1.AddItem ""
    
    dbTabCode.Close
    dbTabCode.Open "SELECT *  FROM TabCode ORDER BY Code", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    
    
END_FN:
    On Error GoTo 0
    Frm.List1.AddItem Now & " - Import Procedure Finished."
    Frm.List1.AddItem ""
    If rc Then
        Dim PATH As String
        Call SplitPathFile(file_name, PATH)
        SaveSetting App.Title, "ImportExcel", "FileName0", file_name
        SaveSetting App.Title, "ImportExcel", "Date0", Now
        SaveSetting App.Title, "ImportExcel", "Path0", PATH
        PopupMessage 2, "Excel Hanna Code Import Procedure Finished...."
    End If
    FileExcel.Close False
    Set FileExcel = Nothing

    HannaCodeExcelImport = rc
    
    Exit Function
    
ERR_CREATE_OBJECT:
   ' MsgBox Err.ProductName
    Frm.List1.AddItem Now & " - " & err.Description
   ' PopupMessage 2, "Excel Import Procedure Failed...." & vbCrLf & Err.ProductName
    Resume Next

End Function


Public Function CopyHannaCodeData(ByRef pBar As ProgressBar) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim x As Integer
Dim strDecimal As String
Dim strFields As String

    On Error GoTo ERR_COPY
    
    FormatPage
    
    rc = True
    i = 0
    
    With dbTabCode
        .Filter = ""
        If .EOF Then
            rc = False
            GoTo ERR_END:
        Else
            .MoveFirst
            pBar.Max = .RecordCount
            x = 1
            For t = 1 To 55 '.fields.Count - 1
                Select Case t
                    Case 54
                         GoTo cont3:
                End Select
                      x = x + 1
                      
                    strFields = Trim(.Fields(t).Name)
                    strFields = Replace(strFields, "STDMR", "MR")
                    
                   Call AddCodeValue(1, x + 2, IIf(IsNull(strFields), "", strFields))
cont3:
            Next
                
            Do
                i = i + 1
                x = 1
                
                If IsNull(Trim(.Fields(2))) Then
                    GoTo cont
                End If
                    
                For t = 1 To 55 '.fields.Count - 1
                
                    If IsNull(Trim(.Fields(t))) Then
                        x = x + 1
                        GoTo cont2
                    End If
                    
                    strFields = Trim(.Fields(t))
                    strFields = Replace(strFields, "STDMR", "MR")
                    
                    Select Case t
                        Case 54
                            x = x + 1
                             GoTo cont2:
                    End Select
                    x = x + 1
                      
                      
                    If i = 54 Then GoTo cont:
                    If t = 17 Or t = 18 Then
                        If InStr(strFields, "/") Then
                            GoTo TrueValue
                       Else
                        Call AddCodeValue(i + 1, x + 2, IIf(IsNull(strFields), "", strFields & "%"))
                        End If
                    ElseIf t = 50 Then
                        If InStr(Trim(UCase(.Fields(t))), "FALSE") Then
                        Else
                            Call AddCodeValue(i + 1, x + 2, IIf(IsNull(strFields), "", strFields))
                        End If
                    Else
TrueValue:
                        Call AddCodeValue(i + 1, x + 2, IIf(IsNull(strFields), "", strFields))
                    End If
cont2:
                Next
cont:
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
    Resume Next
End Function
Public Function DeleteAllTabCode()


    dbCode.Execute "DELETE * FROM TabCode"
    DoEvents
    dbTabCode.Close
    dbTabCode.Open "SELECT *  FROM TabCode ORDER BY Code", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    
    DoEvents
   
End Function
