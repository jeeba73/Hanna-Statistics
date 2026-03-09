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
        PopupMessage 3, "Canno't find data file...", , True, ProductionFileName
        Exit Function
    End If
    
    If bFinalQC Then
        MyQRCode.Text3 = "Final QC"
    Else
        
        MyQRCode.Text3 = "Production QC"
        
        If F_MsgBox.DoShow("Select Final QC or Production QC QRCode!", "QRCode : " & HannaCode, , "Final QC", "Prod QC") Then
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
            MyQRCode.Code = IIf(IsNull(Trim(!HannaCode)), "", Trim(!HannaCode))
            MyQRCode.Lot = IIf(IsNull(Trim(!Lot)), "", Trim(!Lot))
            MyQRCode.Exp = IIf(IsNull(Trim(!ExpDate)), "", Trim(!ExpDate))
            MyQRCode.Operator = MyOperatore.Name
            MyQRCode.Date = FormatDataLAT(FormatDateTime(Now(), vbShortDate))
            MyQRCode.Time = FormatDateTime(Now(), vbShortTime)
            
            

        End If
    
    End With
    
    
       rc = DoPrintLabelProduction(MyQRCode)
       

    If rc Then
        PopupMessage 3, "Printing Label....", HannaCode
    End If

    CloseSettingDataFile
    
End Function

