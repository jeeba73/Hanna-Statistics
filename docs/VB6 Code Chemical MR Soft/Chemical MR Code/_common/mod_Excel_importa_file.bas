Attribute VB_Name = "mod_Excel_importa_file"
Option Explicit


'variabile oggetto che contiene il riferimento alla cartella di lavoro di Excel
Private FileExcel As Object

'variabile oggetto che contiene il riferimento al foglio di lavoro di Excel

'Private FoglioExcel as object
Private FoglioExcel As Object

'variabile oggetto che contiene il  riferimento alle celle del foglio di lavoro
'di Excel
Private CellaFoglioExcel As Range

Private MyMRArray() As MRType
Private MyMRWarehouseArray() As WareHouseEntry

Public Function DeleteAllTabCode()


    dbCode.Execute "DELETE * FROM TabCode"
    DoEvents
    dbTabCode.Close
    dbTabCode.Open "SELECT *  FROM TabCode ORDER BY Code", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    
    DoEvents
   
End Function

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
    Frm.List1.AddItem Now & " - " & Err.ProductName
   ' PopupMessage 2, "Excel Import Procedure Failed...." & vbCrLf & Err.ProductName
    Resume Next

End Function



Public Function CopyHannaCodeData(ByRef pBar As ProgressBar) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim x As Integer
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
            x = 1
            For t = 1 To .fields.Count - 1
                
               ' Select Case t
               '     Case 7, 8, 15, 16, 17, 18, 21, 22, 24, 25, 27, 28, 30, 31, 33, 34, 67
                '         GoTo cont:
               '    Case 36 To 51
               '          GoTo cont:
               ' End Select
                
                x = x + 1
                Call AddCodeValue(1, x + 2, IIf(IsNull(Trim(.fields(t).Name)), "", "'" & Trim(.fields(t).Name)))
cont:
            Next
                
            Do
                i = i + 1
                x = 1
                For t = 1 To .fields.Count - 1
                    
                      x = x + 1

TrueValue:
                        Call AddCodeValue(i + 1, x + 2, IIf(IsNull(Trim(.fields(t))), "", "'" & Trim(.fields(t))))
                    
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
    MsgBox Err.Description
    GoTo ERR_END:
End Function



Public Function CopyChemicalMRData(ByRef pBar As ProgressBar) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim strDecimal As String

    On Error GoTo ERR_COPY
    
    FormatPage
    
    rc = True
    i = 0
    
    With dbTabMR
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
                        
                    'ElseIf t = 14 Or t = 15 Or t = 17 Then
                    ''    strDecimal = IIf(IsNull(.fields(14)), "", .fields(14))
                    '    strDecimal = FormatDecimal(strDecimal)
                    '    Call AddCodeValue(i + 1, t + 2, IIf(IsNull(Trim(.fields(t))), "", Format$(Trim(.fields(t)), strDecimal)))
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
    CopyChemicalMRData = rc
    Exit Function
ERR_COPY:
    rc = False
    MsgBox Err.Description
    GoTo ERR_END:
End Function
Public Function iExcelChemicalMR(ByVal mFile As String, ByVal Frm As Form, Optional ByRef DataCounter As Long, Optional ByVal bDeletePreviousRecords As Boolean) As Boolean
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

If bDeletePreviousRecords Then
   
   Call DeleteAllTabMR
  

End If


MyMRArray = MRCleanArray
MyMRWarehouseArray = MyWareHouseEntryCleanArray

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
            
    Frm.List1.AddItem Now & " - Loading Chemical MR ..."
    
            
    r = 3
    i = 0


    Dim NewCode As Integer
    Dim sString As String
    Do
        i = i + 1
        r = r + 1
        
            
        ReDim Preserve MyMRArray(r)
        ReDim Preserve MyMRWarehouseArray(r)
            
               
        MyMRArray(i).Code = Trim(FoglioExcel.Cells(r, 1))
        MyMRArray(i).Description = Trim(FoglioExcel.Cells(r, 3))
        
        MyMRArray(i).Supplier = Trim(FoglioExcel.Cells(r, 4))
        MyMRArray(i).MNP = Trim(FoglioExcel.Cells(r, 5))
        MyMRArray(i).PhysicalState = Trim(FoglioExcel.Cells(r, 7))
        MyMRArray(i).Density = Trim(FoglioExcel.Cells(r, 8))
        MyMRArray(i).Unit = Trim(FoglioExcel.Cells(r, 11))
        MyMRArray(i).Parameter = Trim(FoglioExcel.Cells(r, 12))
        MyMRArray(i).FWParameter = Trim(FoglioExcel.Cells(r, 13))
        MyMRArray(i).StorageT = Trim(FoglioExcel.Cells(r, 15))
        MyMRArray(i).MinQty = Trim(FoglioExcel.Cells(r, 18))
 
 
        ReDim MyMRWarehouseArray(i).Bottle(0)
        
        MyMRWarehouseArray(i).MRCode = Trim(FoglioExcel.Cells(r, 1))
        MyMRWarehouseArray(i).Description = Trim(FoglioExcel.Cells(r, 3))
        MyMRWarehouseArray(i).Bottle(0) = Trim(FoglioExcel.Cells(r, 2))
        MyMRWarehouseArray(i).Lot = Trim(FoglioExcel.Cells(r, 6))
        MyMRWarehouseArray(i).Purity = Trim(FoglioExcel.Cells(r, 9))
        Debug.Print Trim(FoglioExcel.Cells(r, 10))
        MyMRWarehouseArray(i).MRValueConcentration = Trim(FoglioExcel.Cells(r, 10))
        MyMRWarehouseArray(i).Location = Trim(FoglioExcel.Cells(r, 14))
        
        
        MyMRWarehouseArray(i).StockQTY = Trim(FoglioExcel.Cells(r, 16))
        MyMRWarehouseArray(i).stockUnit = Trim(FoglioExcel.Cells(r, 17))
        MyMRWarehouseArray(i).ArrivedTime = Trim(FoglioExcel.Cells(r, 19))
        MyMRWarehouseArray(i).Open = Trim(FoglioExcel.Cells(r, 10))
        MyMRWarehouseArray(i).Finished = Trim(FoglioExcel.Cells(r, 21))
        MyMRWarehouseArray(i).SupplierEXP = Trim(FoglioExcel.Cells(r, 22))
        MyMRWarehouseArray(i).MREXP = Trim(FoglioExcel.Cells(r, 23))
        MyMRWarehouseArray(i).Status = Trim(FoglioExcel.Cells(r, 24))
        MyMRWarehouseArray(i).Note = Trim(FoglioExcel.Cells(r, 25))

        If MyMRArray(i).Code = "" Then
            GoTo END_FN
        End If
        
        With dbTabMR
            .filter = ""
            .filter = "Code='" & MyMRArray(i).Code & "'"
            If Not (.EOF) Then
                
                ' ok controllo che sia associata al cliente
                Frm.List1.AddItem Now & " - Chemical MR  (" & i & ") : " & MyMRArray(i).Code & " ( " & MyMRArray(i).Description & " ) already Exsists... "
            Else
                .AddNew
                NewCode = NewCode + 1
                Frm.List1.AddItem Now & " - Import new Chemical MR  (" & i & ") : " & MyMRArray(i).Code & " ( " & MyMRArray(i).Description & " )"
            End If
                
               !Code = MyMRArray(i).Code
               !Description = MyMRArray(i).Description
               !Supplier = MyMRArray(i).Supplier
               !MNP = MyMRArray(i).MNP
               !PhysicalState = MyMRArray(i).PhysicalState
               !Density = MyMRArray(i).Density
               !Unit = MyMRArray(i).Unit
               !Parameter = MyMRArray(i).Parameter
               !FWParameter = MyMRArray(i).FWParameter
               !StorageT = MyMRArray(i).StorageT
               !MinQty = MyMRArray(i).MinQty
               !STOCK_QTY = 0
               !STOCK_UNIT = MyMRWarehouseArray(i).stockUnit
               !ReductionExpDays = 120
               !Modified = Now()
                .Update
        End With
        
      '  With dbTabMRWarehouse
      '      .filter = ""
      '     ' If MyMRWarehouseArray(i).Bottle <> "" And MyMRWarehouseArray(i).Lot <> "" Then
      '          .filter = "Code='" & MyMRWarehouseArray(i).MRCode & "' and bottle='" & MyMRWarehouseArray(i).Bottle(0) & "' and Lot='" & MyMRWarehouseArray(i).Lot & "'"
      '     ' Else
      '
      '        ' GoTo contadd:
      '    ' End If
      '      If Not (.EOF) Then
      '
      '          ' ok controllo che sia associata al cliente
      '          Frm.List1.AddItem Now & " - Warehouse Entry  (" & i & ") : " & MyMRWarehouseArray(i).MRCode & " ( " & MyMRWarehouseArray(i).Description & " ) already Exsists... "
      '      Else
contadd:
      '          .AddNew
      '          'NewCode = NewCode + 1
      '          Frm.List1.AddItem Now & " - Import new Warehouse Entry (" & i & ") : " & MyMRWarehouseArray(i).MRCode & " ( " & MyMRWarehouseArray(i).Description & " )"
      '      End If
      '
      '
      '          !Code = MyMRWarehouseArray(i).MRCode
      '
      '          !Description = MyMRWarehouseArray(i).Description
      '          !Bottle = MyMRWarehouseArray(i).Bottle(0)
      '          !Lot = MyMRWarehouseArray(i).Lot
      '          !Density = MyMRArray(i).Density
      '          !Purity = MyMRWarehouseArray(i).Purity
      '          !MRValue = MyMRWarehouseArray(i).MRValueConcentration
      '          !Unit = MyMRArray(i).Unit
      '          !Parameter = MyMRArray(i).Parameter
      '          !FWParameter = MyMRArray(i).FWParameter
      '
      '          !Location = MyMRWarehouseArray(i).Location
      '          !StockQTY = IIf(IsNumeric(MyMRWarehouseArray(i).StockQTY), MyMRWarehouseArray(i).StockQTY, 0)
      '          !stockUnit = MyMRWarehouseArray(i).stockUnit
      '
      '          If IsDate(MyMRWarehouseArray(i).ArrivedTime) Then
      '              !ArrivedTime = MyMRWarehouseArray(i).ArrivedTime
      '          End If
      '          If IsDate(MyMRWarehouseArray(i).Open) Then
      '              !Open = MyMRWarehouseArray(i).Open
      '          End If
      '          If IsDate(MyMRWarehouseArray(i).Finished) Then
      '              !Finished = MyMRWarehouseArray(i).Finished
      '          End If
      '          If IsDate(MyMRWarehouseArray(i).SupplierEXP) Then
      '              !SupplierEXP = MyMRWarehouseArray(i).SupplierEXP
      '          End If
      '          If IsDate(MyMRWarehouseArray(i).MREXP) Then
      '              !MREXP = MyMRWarehouseArray(i).MREXP
      '          End If
                
      '          !Status = (MyMRWarehouseArray(i).Status)
      '          !bClosed = IIf(!Status = "2", True, False)
      '          !Note = MyMRWarehouseArray(i).Note
                
 
      '          Frm.List1.AddItem Now & " - Chemical MR  (" & i & ") : " & MyMRArray(i).Code & " ( " & MyMRArray(i).Description & " ) Saved... "
      '
      '          .Update
      '  End With
        
        With dbTabLocation
            .filter = ""
            .filter = "Code='" & MyMRWarehouseArray(i).Location & "'"
            If .EOF Then
                .AddNew
                !Code = MyMRWarehouseArray(i).Location
                .Update
            End If
        End With
        
        
        
    Loop Until (FoglioExcel.Cells(r + 1, 1) = "" And FoglioExcel.Cells(r + 2, 1) = "")

    
    DataCounter = i
    Frm.List1.AddItem ""
    Frm.List1.AddItem "n." & NewCode & " New Chemical MR Records Imported"
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
        PopupMessage 2, "Excel Chemical WareHouse Import Procedure Finished...."
        PopupMessage 2, "Excel Chemical MR Import Procedure Finished...."
    End If
    FileExcel.Close False
    Set FileExcel = Nothing

    iExcelChemicalMR = rc
    
    Exit Function
    
ERR_CREATE_OBJECT:
    'MsgBox Err.Description
    Frm.List1.AddItem Now & " - " & Err.Description
   ' PopupMessage 2, "Excel Import Procedure Failed...." & vbCrLf & Err.Description
    Resume Next
     Exit Function
     
End Function

