Attribute VB_Name = "mod_Excel_ProductClassification"
Option Explicit

'variabile oggetto che contiene il riferimento alla cartella di lavoro di Excel
Private FileExcel As Object

'variabile oggetto che contiene il riferimento al foglio di lavoro di Excel
Private FoglioExcel As Object

'variabile oggetto che contiene il  riferimento alle celle del foglio di lavoro
'di Excel
Private CellaFoglioExcel As Range


Public Function iExcelProductClassification(ByVal mFile As String, ByVal Frm As Form, Optional ByRef DataCounter As Long, Optional ByVal bDeletePreviousRecords As Boolean) As Boolean
Dim rc As Boolean
Dim file_name As String
Dim ClienteID As Long
Dim SedeID As Long
Dim NewID As Long
Dim ImportRangeID As Long
Dim NumRange As Integer
Dim bAssociaCliente As Boolean
Dim strMeter As String

Dim i As Integer
Dim t As Integer
Dim x As Integer
Dim r As Long
Dim nMax As Integer
Dim strDecimal As String

Dim divisioni As Long
Dim Classe As String

Dim sDestinazione As String


MyImportProductClassification = MyImportProductClassificationClean


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
            
    r = 2
    i = 0
    
    
    Do
        r = r + 1
    Loop Until (FoglioExcel.Cells(r + 1, 1) = "" And FoglioExcel.Cells(r + 2, 1) = "")
    
    ReDim MyImportProductClassification(r)
    
    
    Frm.List1.AddItem Now & " - Loading Hanna SFG Code Classification ..."
    
            
    r = 1
    i = 0


    Dim NewCode As Integer
    Dim sString As String
    Do
        i = i + 1
        r = r + 1
        
               
        MyImportProductClassification(i).Code = Trim(FoglioExcel.Cells(r, 1))
        MyImportProductClassification(i).Name = Trim(FoglioExcel.Cells(r, 2))

        If MyImportProductClassification(i).Code = "" Then
            GoTo END_FN
        End If
        
        With dbTabCodeClassification
            .filter = ""
            .filter = "Code='" & MyImportProductClassification(i).Code & "'" ' and RangeMin='" & MyImportProductClassification(i).RangeMin & "' and RangeMax='" & MyImportProductClassification(i).RangeMax & "'"
            If Not (.EOF) Then
                
                ' ok controllo che sia associata al cliente
                Frm.List1.AddItem Now & " - Hanna SFG Code Classification (" & i & ") : " & MyImportProductClassification(i).Code & " ( " & MyImportProductClassification(i).Name & " ) already Exsists... "
            Else
                .AddNew
                NewCode = NewCode + 1
                Frm.List1.AddItem Now & " - Import new Hanna SFG Code Classification (" & i & ") : " & MyImportProductClassification(i).Code & " ( " & MyImportProductClassification(i).Name & " )"
            End If
                
                
                
              
                For t = 1 To .fields.Count - 2
                    
                    sString = Trim(FoglioExcel.Cells(r, t))
                    If sString <> "" Then .fields(t) = Trim(sString)

                Next
                
                !DateModified = Now
                Frm.List1.AddItem Now & " - Hanna SFG Code  Classification (" & i & ") : " & MyImportProductClassification(i).Code & " ( " & MyImportProductClassification(i).Name & " ) Saved... "
                DoEvents
                .Update
        End With
        
        
    Loop Until (FoglioExcel.Cells(r + 1, 1) = "" And FoglioExcel.Cells(r + 2, 1) = "")
    
    
    
    
    
    DataCounter = i
    Frm.List1.AddItem ""
    Frm.List1.AddItem "n." & NewCode & " New Product Classification Records Imported"
    Frm.List1.AddItem "n." & DataCounter & " Excel Code"
    Frm.List1.AddItem ""
    

    
END_FN:
    On Error GoTo 0
    Frm.List1.AddItem Now & " - Import Procedure Finished."
    Frm.List1.AddItem ""
    If rc Then
        Dim PATH As String
        Call SplitPathFile(file_name, PATH)
        SaveSetting App.Title, "ImportExcel", "FileName2", file_name
        SaveSetting App.Title, "ImportExcel", "Date2", Now
        SaveSetting App.Title, "ImportExcel", "Path2", PATH
        PopupMessage 2, "Excel Product Classification Import Procedure Finished...."
    End If
    FileExcel.Close False
    Set FileExcel = Nothing

    iExcelProductClassification = rc
    
    Exit Function
    
ERR_CREATE_OBJECT:
    MsgBox Err.Description
    Frm.List1.AddItem Now & " - " & Err.Description
    PopupMessage 2, "Excel Import Procedure Failed...." & vbCrLf & Err.Description
    Resume Next
     Exit Function
     
End Function


Public Function iExcelFrasiH(ByVal mFile As String, ByVal Frm As Form, Optional ByRef DataCounter As Long, Optional ByVal bDeletePreviousRecords As Boolean) As Boolean
Dim rc As Boolean
Dim file_name As String
Dim ClienteID As Long
Dim SedeID As Long
Dim NewID As Long
Dim ImportRangeID As Long
Dim NumRange As Integer
Dim bAssociaCliente As Boolean
Dim strMeter As String

Dim i As Integer
Dim t As Integer
Dim x As Integer
Dim r As Long
Dim nMax As Integer
Dim strDecimal As String

Dim divisioni As Long
Dim Classe As String

Dim sDestinazione As String


MyFrasiH = MyFrasiHClean

If bDeletePreviousRecords Then
    
    Call DeleteAllTabFrasiH
   

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
            
    r = 2
    i = 0
    
    
    Do
        r = r + 1
    Loop Until (FoglioExcel.Cells(r + 1, 3) = "" And FoglioExcel.Cells(r + 2, 3) = "")
    
    ReDim MyFrasiH(r)
    
    
    Frm.List1.AddItem Now & " - Loading Phrases H Code  ..."
    
            
    r = 2
    i = 0


    Dim NewCode As Integer
    Dim sString As String
    Do
        i = i + 1
        r = r + 1
        
               
        MyFrasiH(i).Code = Trim(FoglioExcel.Cells(r, 3))
        MyFrasiH(i).PhyHazStatement = Trim(FoglioExcel.Cells(r, 4))

        If MyFrasiH(i).Code = "" Then
            GoTo END_FN
        End If
        
        With dbTabFrasiH
            .filter = ""
            .filter = "Code='" & MyFrasiH(i).Code & "'" ' and RangeMin='" & MyFrasiH(i).RangeMin & "' and RangeMax='" & MyFrasiH(i).RangeMax & "'"
            If Not (.EOF) Then
                
                ' ok controllo che sia associata al cliente
                Frm.List1.AddItem Now & " - Phrases H Code  (" & i & ") : " & MyFrasiH(i).Code & " ( " & MyFrasiH(i).PhyHazStatement & " ) already Exsists... "
            Else
                .AddNew
                NewCode = NewCode + 1
                Frm.List1.AddItem Now & " - Import new Phrases H Code (" & i & ") : " & MyFrasiH(i).Code & " ( " & MyFrasiH(i).PhyHazStatement & " )"
            End If
                
                
                
              
                For t = 1 To .fields.Count - 2
                    
                    sString = Trim(FoglioExcel.Cells(r, t + 2))
                    If sString <> "" Then .fields(t) = Trim(sString)

                Next
                
                !DateModified = Now
                Frm.List1.AddItem Now & " - Hanna Phrases H Code (" & i & ") : " & MyFrasiH(i).Code & " ( " & MyFrasiH(i).PhyHazStatement & " ) Saved... "
                DoEvents
                .Update
        End With
        
        
    Loop Until (FoglioExcel.Cells(r + 1, 3) = "" And FoglioExcel.Cells(r + 2, 3) = "")
    
    
    
    
    
    DataCounter = i
    Frm.List1.AddItem ""
    Frm.List1.AddItem "n." & NewCode & " New Phrases H Records Imported"
    Frm.List1.AddItem "n." & DataCounter & " Excel Code"
    Frm.List1.AddItem ""
    

    
END_FN:
    On Error GoTo 0
    Frm.List1.AddItem Now & " - Import Procedure Finished."
    Frm.List1.AddItem ""
    If rc Then
        Dim PATH As String
        Call SplitPathFile(file_name, PATH)
        SaveSetting App.Title, "ImportExcel", "FileName3", file_name
        SaveSetting App.Title, "ImportExcel", "Date3", Now
        SaveSetting App.Title, "ImportExcel", "Path3", PATH
        PopupMessage 2, "Excel Product Classification Import Procedure Finished...."
    End If
    FileExcel.Close False
    Set FileExcel = Nothing

    iExcelFrasiH = rc
    
    Exit Function
    
ERR_CREATE_OBJECT:
   ' MsgBox err.Description
    Frm.List1.AddItem Now & " - " & Err.Description
    PopupMessage 2, "Excel Import Procedure Failed...." & vbCrLf & Err.Description
    GoTo END_FN
     Exit Function
     
End Function

