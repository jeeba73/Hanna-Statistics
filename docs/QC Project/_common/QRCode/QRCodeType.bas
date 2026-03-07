Attribute VB_Name = "QRCodeTypeReadings"
Option Explicit
Private SettingName As String



Public Function PrintLabelReadings(ByVal Lot As String, ByVal code As String)
Dim MyQRCode As Qr01
Dim PREP_PATH As String
Dim HannaCode As String
Dim MaxHannaCode As String
Dim i As Integer
Dim rc As Boolean
Dim stringNote As String
Dim m_rc As Boolean
Dim RfPfileName As String


    CloseSettingDataFile
    
    rc = False
    

    
     If MyOperatore.Name = "" Then

    If frmLogin.DoShow Then
            
    Else
        Exit Function
    End If

    End If

    
    MyQRCode.QC = "Failed"
    
    If F_MsgBox.DoShow("Enter QC Evaluation", "QC : " & code, , "Passed", "Failed") Then MyQRCode.QC = "Passed"
  
    If F_InputBox.DoShow("Enter QC Note", "QRCode : " & code, , , , stringNote) Then MyQRCode.Note = stringNote
     
    With dbTabReport
        .filter = ""
        .filter = "Code='" & code & "' and Lot='" & Lot & "'"
        If .EOF Then
            '--------------------------------
            ' dovrei cercare nei file
            '-------------------------------
            
        Else
            MyQRCode.Recipe = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
            MyQRCode.code = IIf(IsNull(Trim(!code)), "", Trim(!code))
            MyQRCode.Exp = IIf(IsNull(Trim(!Exp)), "", Trim(!Exp))
            MyQRCode.Lot = IIf(IsNull(Trim(!Lot)), "", Trim(!Lot))
            MyQRCode.date = FormatDataLAT(FormatDateTime(Now(), vbShortDate))
            MyQRCode.Time = FormatDateTime(Now(), vbShortTime)
            MyQRCode.Operator = MyOperatore.Name
            MyQRCode.Line = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
        End If
    
    End With

    rc = DoPrintLabelReadings(MyQRCode, False)

        
    If rc Then
        PopupMessage 2, "Printing Label....", MyQRCode.code
    End If

    CloseSettingDataFile
    
End Function


Public Function SetQrCodePreparation(ByRef QRCode As Qr01, ByVal PreparationFileName As String, ByVal PreparationID As String, ByVal i As Integer) As Boolean


Dim rc As Boolean
Dim PREP_PATH As String
Dim strCode As String

    On Error GoTo ERR_SET:
    
    rc = True
    
    SettingName = PreparationFileName
    
    With QRCode
    strCode = GetSettingData(SettingName, "HannaCode" & i, "Code", "", PREP_PATH)
      .code = IIf(strCode = "", .code, strCode)
      .Qty = GetSettingData(SettingName, "HannaCode" & i, "QtyToProduce", "", PREP_PATH) & " " & GetSettingData(SettingName, "HannaCode" & i, "Um", "", PREP_PATH)
    End With
    
    CloseSettingDataFile
  
ERR_END:
    On Error GoTo 0
    CloseSettingDataFile
    SetQrCodePreparation = rc
    Exit Function
ERR_SET:
    rc = False
    Resume Next

End Function

