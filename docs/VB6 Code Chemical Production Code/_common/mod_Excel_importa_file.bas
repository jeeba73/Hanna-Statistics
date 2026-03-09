Attribute VB_Name = "mod_Excel_Importa_files"
Option Explicit


'variabile oggetto che contiene il riferimento alla cartella di lavoro di Excel
Private FileExcel As Object

'variabile oggetto che contiene il riferimento al foglio di lavoro di Excel
Private FoglioExcel As Object

'variabile oggetto che contiene il  riferimento alle celle del foglio di lavoro
'di Excel
Private CellaFoglioExcel As Range

Public MyImportHannaCode As HannaCode



Public Function HannaCodeExcelImport(ByVal mFile As String, ByVal Frm As Form, Optional ByRef CodeCount As Long, Optional ByVal bDeletePreviousRecords As Boolean) As Boolean
Dim rc As Boolean
Dim file_name As String


Dim i As Integer
Dim t As Integer
Dim X As Integer
Dim r As Long
Dim nMax As Integer
Dim strRecipe As String


Dim sDestinazione As String


If bDeletePreviousRecords Then
   
   Call DeleteAllTabCode
   
   If F_MsgBox.DoShow("Hanna Code Table correctly deleted. Delete Recipes Table Too?", "Import Hanna Codes", , "Delete", "Don't") Then
     
     Call DeleteAllTabRecipe
   
   End If

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

               
        MyImportHannaCode.Code = Trim(FoglioExcel.Cells(r, 2))
        MyImportHannaCode.ProductName = Trim(FoglioExcel.Cells(r, 5))
        MyImportHannaCode.RangeMin = Trim(FoglioExcel.Cells(r, 30))
        MyImportHannaCode.RangeMax = Trim(FoglioExcel.Cells(r, 31))
        
      
        
        If MyImportHannaCode.Code = "" Then
            GoTo END_FN
        End If
        
        With dbTabCode
            .filter = ""
            
            If MyImportHannaCode.RangeMin = "" Or MyImportHannaCode.RangeMax = "" Then
                 .filter = "Code='" & MyImportHannaCode.Code & "'"
            Else
                .filter = "Code='" & MyImportHannaCode.Code & "' and RangeMin='" & MyImportHannaCode.RangeMin & "' and RangeMax='" & MyImportHannaCode.RangeMax & "'"
            End If
            If Not (.EOF) Then
                
                ' ok controllo che sia associata al cliente
                Frm.List1.AddItem Now & " - Hanna SFG Code (" & i & ") : " & MyImportHannaCode.Code & " ( " & MyImportHannaCode.ProductName & " ) already Exsists... "
            Else
                .AddNew
                NewCode = NewCode + 1
                Frm.List1.AddItem Now & " - Import new Hanna SFG Code (" & i & ") : " & MyImportHannaCode.Code & " ( " & MyImportHannaCode.ProductName & " )"
            End If
                
            For t = 1 To .fields.Count - 1

                   If InStr(.fields(t).Name, "Date") Then
                     !DateModified = Now
                  ' ElseIf InStr(.fields(t).Name, "Hide") Then
                   ' !Hide = False
                   Else
                   
                    .fields(t).Value = Trim(FoglioExcel.Cells(r, t + 1))
                End If
       
            Next
            
            
               
                Frm.List1.AddItem Now & " - Hanna SFG Code (" & i & ") : " & MyImportHannaCode.Code & " ( " & MyImportHannaCode.ProductName & " ) Saved... "
                DoEvents
                .Update
                
              
        End With
    Loop Until (FoglioExcel.Cells(r + 1, 2) = "" And FoglioExcel.Cells(r + 2, 2) = "")
    
    CodeCount = i
    Frm.List1.AddItem ""
    Frm.List1.AddItem "n." & NewCode & " New Hanna Code Imported"
    Frm.List1.AddItem "n." & CodeCount & " Excel Code"
    Frm.List1.AddItem ""
    
    
END_FN:
    On Error GoTo 0
    Frm.List1.AddItem Now & " - Import Procedure Finished."
    Frm.List1.AddItem ""
    If rc Then
        Dim Path As String
        Call SplitPathFile(file_name, Path)
        SaveSetting App.Title, "ImportExcel", "FileName0", file_name
        SaveSetting App.Title, "ImportExcel", "Date0", Now
        SaveSetting App.Title, "ImportExcel", "Path0", Path
        PopupMessage 2, "Excel Hanna Code Import Procedure Finished...."
    End If
    FileExcel.Close False
    Set FileExcel = Nothing

    HannaCodeExcelImport = rc
    
    Exit Function
    
ERR_CREATE_OBJECT:
   ' MsgBox Err.ProductName
    Frm.List1.AddItem Now & " - " & err.ProductName
   ' PopupMessage 2, "Excel Import Procedure Failed...." & vbCrLf & Err.ProductName
    Resume Next

End Function


Public Function iExcelChemicalRM(ByVal mFile As String, ByVal Frm As Form, Optional ByRef DataCounter As Long, Optional ByVal bDeletePreviousRecords As Boolean) As Boolean
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
Dim X As Integer
Dim r As Long
Dim nMax As Integer
Dim strDecimal As String

Dim divisioni As Long
Dim Classe As String

Dim sDestinazione As String


MyRawMaterial = MyRawMaterialClean


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
    Set FoglioExcel = FileExcel.Worksheets(3)
            
    r = 4
    i = 0
    
    
    Do
        r = r + 1
    Loop Until (FoglioExcel.Cells(r + 1, 1) = "" And FoglioExcel.Cells(r + 2, 1) = "")
    
    ReDim MyRawMaterial(r)
    
    
    Frm.List1.AddItem Now & " - Loading Chemical RM ..."
    
            
    r = 4
    i = 0


    Dim NewCode As Integer
    Dim sString As String
    Do
        i = i + 1
        r = r + 1
        
               
        MyRawMaterial(i).Code = Trim(FoglioExcel.Cells(r, 1))
        MyRawMaterial(i).Description = Trim(FoglioExcel.Cells(r, 2))

        If MyRawMaterial(i).Code = "" Then
            GoTo END_FN
        End If
        
        With dbTabRawMaterial
            .filter = ""
            .filter = "Code='" & MyRawMaterial(i).Code & "'" ' and RangeMin='" & MyRawMaterial(i).RangeMin & "' and RangeMax='" & MyRawMaterial(i).RangeMax & "'"
            If Not (.EOF) Then
                
                ' ok controllo che sia associata al cliente
                Frm.List1.AddItem Now & " - Chemical RM  (" & i & ") : " & MyRawMaterial(i).Code & " ( " & MyRawMaterial(i).Description & " ) already Exsists... "
            Else
                .AddNew
                NewCode = NewCode + 1
                Frm.List1.AddItem Now & " - Import new Chemical RM  (" & i & ") : " & MyRawMaterial(i).Code & " ( " & MyRawMaterial(i).Description & " )"
            End If
                
                
                
              
                For t = 1 To .fields.Count - 3
                    

                    sString = Replace(Trim(FoglioExcel.Cells(r, t)), Chr$(10), "")
                    sString = Replace(Trim(FoglioExcel.Cells(r, t)), Chr$(13), "")
                    
                    If .fields(t).Name = "Um" Then
                        
                        sString = LCase(Replace(UCase(sString), "GR", "g"))
                        ' DA CONFEMRARE
                        If sString = "" Then sString = "g"
                        
                        .fields(t) = Trim(sString)
                    Else
                    
                        .fields(t) = Trim(Left(sString, 255))
                        
                        

                    End If

                Next
                
                If IsNull(!ManufacturerName) Or !ManufacturerName = "" Then
                    !bMix = False
                Else
                    !bMix = IIf(InStr(!ManufacturerName, "Hanna"), True, False)
                End If
                
                !DateModified = Now
                
                Frm.List1.AddItem Now & " - Chemical RM  (" & i & ") : " & MyRawMaterial(i).Code & " ( " & MyRawMaterial(i).Description & " ) Saved... "
            
                .Update
        End With
        
        
    Loop Until (FoglioExcel.Cells(r + 1, 1) = "" And FoglioExcel.Cells(r + 2, 1) = "")
    
    
    
    
    
    DataCounter = i
    Frm.List1.AddItem ""
    Frm.List1.AddItem "n." & NewCode & " New Chemical RM Records Imported"
    Frm.List1.AddItem "n." & DataCounter & " Excel Code"
    Frm.List1.AddItem ""
    

    
END_FN:
    On Error GoTo 0
    Frm.List1.AddItem Now & " - Import Procedure Finished."
    Frm.List1.AddItem ""
    If rc Then
        Dim Path As String
        Call SplitPathFile(file_name, Path)
        SaveSetting App.Title, "ImportExcel", "FileName2", file_name
        SaveSetting App.Title, "ImportExcel", "Date2", Now
        SaveSetting App.Title, "ImportExcel", "Path2", Path
        PopupMessage 2, "Excel Chemical RM Import Procedure Finished...."
    End If
    FileExcel.Close False
    Set FileExcel = Nothing

    iExcelChemicalRM = rc
    
    Exit Function
    
ERR_CREATE_OBJECT:
    'MsgBox Err.Description
    Frm.List1.AddItem Now & " - " & err.Description
   ' PopupMessage 2, "Excel Import Procedure Failed...." & vbCrLf & Err.Description
    Resume Next
     Exit Function
     
End Function




Private Function CheckRecipesFromHannaCodes(ByVal sString As String, ByVal strMix1 As String, ByVal strMix2 As String, ByVal strLine As String, ByVal strExp As String, ByVal strRev As String) As Boolean
Dim rc As Boolean
Dim strMix As String
    rc = True
    On Error GoTo ERR_CHECK:
    
'------------------------------------------------------------
'
'   da Import Excel di Hanna Code ho le Ricette e i Mix
'   se non sono presenti nel DB Recipes allora li inserisco
'
'-------------------------------------------------------------
    
    strMix = ""
    If strMix1 = "" Then
    Else
    
        If sString = strMix1 Or sString = strMix2 Then
        
        Else
        
            ' aggiungo il Mix al RawMaterial!!!
            Call AddMixToRawMaterial(strMix1)
            
        End If
    
    
        strMix = strMix1
    End If
    
    
    
    If strMix2 = "" Then
    Else
        Call AddMixToRawMaterial(strMix2)
        strMix = Trim(IIf(strMix = "", "", strMix & ";") & strMix2)
    End If
        
    With dbTabRecipe
        .filter = ""
        .filter = "Code='" & sString & "'"
        If .EOF Then
            .AddNew
            !Code = sString

        Else
            !Rev = strRev
            !Exp = strExp
            !Line = strLine
            
            If Trim(UCase(sString)) = Trim(UCase(strMix)) Then
            Else
                !Mix = strMix
            End If
        End If
    
        .Update
    End With
    
    
ERR_END:
    On Error GoTo 0
    CheckRecipesFromHannaCodes = rc
    Exit Function
ERR_CHECK:
    rc = False
    GoTo ERR_END
End Function


Private Function AddMixToRawMaterial(ByVal sMix As String)

'------------------------------------------------------------
'
'   Se è un Mix lo inserisco anche come RaW mATERIAL
'
'-------------------------------------------------------------

    With dbTabRawMaterial
        .filter = ""
        .filter = "Code='" & sMix & "'"
        If .EOF Then
            .AddNew
            !Code = sMix
          
            !bMix = True
            .Update
        End If
    End With
    
    
End Function

