Attribute VB_Name = "QRCodeTypeProduction"


Public Function PrintLabelProduction(ByVal ProductionFileName As String, ByVal ProductionID As Long, ByVal HannaCode As String, Optional ByVal bFinalQC As Boolean)
Dim MyQRCode As Qr01
Dim PREP_PATH As String
Dim MaxHannaCode As Integer
Dim i As Integer
Dim rc As Boolean
Dim stringNote As String

    'ProductionID
    'ProductionCode
    'HannaCode
    'ProductionFileName
    'ProductionQC
    
    If MyOperatore.Name = "" Then

    If frmLogin.DoShow Then
            
    Else
        Exit Function
    End If

    End If

    
    CloseSettingDataFile
    
    rc = True
    
    If FileExists(USER_PRODUCTION_PATH & ProductionFileName) Then
        PREP_PATH = USER_PRODUCTION_PATH
    ElseIf FileExists(USER_PRODUCTION_PATH & "data\" & ProductionFileName) Then
        PREP_PATH = USER_PRODUCTION_PATH & "data\"
    Else
        rc = False
        PopupMessage 2, "Canno't find data file...", , True, ProductionFileName
        Exit Function
    End If
    
    If bFinalQC Then
        MyQRCode.Text3 = "Production to QC - P/Final"
    Else
        
        MyQRCode.Text3 = "Production to QC - P/Prod"
        
        If F_MsgBox.DoShow("Select Final QC or Production QC!", "QRCode : " & HannaCode, , "QC - P/Final", "QC - P/Prod") Then
            MyQRCode.Text3 = "Final QC"
        End If
        
    End If

  
    If F_InputBox.DoShow("Enter Production Note", "QRCode : " & HannaCode, , , , stringNote) Then MyQRCode.Note = stringNote
    
    
     
    With dbTabProduction
        .filter = ""
        .filter = "ID='" & ProductionID & "'"
        If .EOF Then
            rc = False
            Exit Function
        Else
            MyQRCode.Recipe = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
            MyQRCode.Code = HannaCode 'IIf(IsNull(Trim(!HannaCode)), "", Trim(!HannaCode))
            MyQRCode.Lot = IIf(IsNull(Trim(!Lot)), "", Trim(!Lot))
            MyQRCode.Exp = IIf(IsNull(Trim(!ExpDate)), "", Trim(!ExpDate))
            MyQRCode.Operator = MyOperatore.Name
            MyQRCode.Date = FormatDataLAT(FormatDateTime(Now(), vbShortDate))
            MyQRCode.Time = FormatDateTime(Now(), vbShortTime)
            MyQRCode.Line = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
        End If
    
    End With
    
     If InStr(MyQRCode.Recipe, ";") Then
        Ricette = Split(MyQRCode.Recipe, ";")
        MyQRCode.Recipe = Trim(Ricette(0))
    End If
    
       rc = DoPrintLabelProduction(MyQRCode)
       

    If rc Then
        PopupMessage 2, "Printing Label....", HannaCode
    End If

    CloseSettingDataFile
    
End Function



Public Function PrintLabelCloseProduction(ByVal ProductionFileName As String, ByVal ProductionID As Long, ByVal HannaCode As String, ByVal lRowHanna As Integer)
Dim MyQRCode As Qr01
Dim PROD_PATH As String
Dim MaxHannaCode As Integer
Dim i As Integer
Dim rc As Boolean
Dim stringNote As String
Dim MaxBoxNumber As Integer
Dim MyUM As String
Dim Ricette As Variant

    'ProductionID
    'ProductionCode
    'HannaCode
    'ProductionFileName
    'ProductionQC
    
    If MyOperatore.Name = "" Then

    If frmLogin.DoShow Then
            
    Else
        Exit Function
    End If

    End If

    
    CloseSettingDataFile
    
    rc = True
    
     If FileExists(USER_PRODUCTION_PATH & ProductionFileName) Then
        PROD_PATH = USER_PRODUCTION_PATH
    ElseIf FileExists(USER_PRODUCTION_PATH & "data\" & ProductionFileName) Then
        PROD_PATH = USER_PRODUCTION_PATH & "data\"
    Else
        rc = False
        PopupMessage 2, "Canno't find data file...", , True, ProductionFileName
        Exit Function
    End If
    
  
    If F_InputBox.DoShow("Enter Production Note", "QRCode : " & HannaCode, , , , stringNote) Then MyQRCode.Note = stringNote
    
    stringNote = 1
    If F_InputBox.DoShow("Enter Max BOX Number", "#Box : " & HannaCode, , , , stringNote, , True) Then MaxBoxNumber = CInt(stringNote)
     
    With dbTabProduction
        .filter = ""
        .filter = "ID='" & ProductionID & "'"
        If .EOF Then
            rc = False
            Exit Function
        Else
            MyQRCode.Recipe = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
            MyQRCode.Code = HannaCode 'IIf(IsNull(Trim(!HannaCode)), "", Trim(!HannaCode))
            MyQRCode.Lot = IIf(IsNull(Trim(!Lot)), "", Trim(!Lot))
            MyQRCode.Exp = IIf(IsNull(Trim(!ExpDate)), "", Trim(!ExpDate))
            MyQRCode.Operator = MyOperatore.Name
            MyQRCode.Date = FormatDataLAT(FormatDateTime(Now(), vbShortDate))
            MyQRCode.Time = FormatDateTime(Now(), vbShortTime)
            
            

        End If
    
    End With
    
    ' non chiedo più il maxboxnumber
    ' MaxBoxNumber = 1
    
    If InStr(MyQRCode.Recipe, ";") Then
        Ricette = Split(MyQRCode.Recipe, ";")
        MyQRCode.Recipe = Trim(Ricette(0))
    End If
    
     lRowHanna = IIf(lRowHanna = 0, 1, lRowHanna)
     
    MyUM = GetSettingData(ProductionFileName, "HannaCode" & lRowHanna, "Um", "", PROD_PATH)
    MyQRCode.TotalQty = GetSettingData(ProductionFileName, "HannaCode" & lRowHanna, "QtyProduced", "0", PROD_PATH)
    
    CloseSettingDataFile
    
   
     stringNote = ""
        If F_InputBox.DoShow("Enter BOX Quantity", "BOX QTY", , , , stringNote, , True) Then
        MyQRCode.Qty = (stringNote) '& " " & myUm
       End If
       
    For i = 1 To MaxBoxNumber
    
        'stringNote = GetSetting(App.Title, "Production", "BoxN", 1)

        'If F_InputBox.DoShow("Enter BOX Number", "#Box : " & HannaCode, , , , stringNote, , True) Then
            ' remember numbr of Box
            
           ' SaveSetting App.Title, "Production", "BoxN", stringNote

        MyQRCode.Box = i ' CInt(stringNote) ' & "/" & MaxBoxNumber
        'End If
       
       rc = DoPrintLabelCloseProduction(MyQRCode)
       



    Next
       
    If F_MsgBox.DoShow("Print final Quantity label?", "Close Production") Then
            Call DoPrintLabelCloseProduction(MyQRCode, True)
    End If

    CloseSettingDataFile
    
End Function


