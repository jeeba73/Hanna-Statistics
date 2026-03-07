Attribute VB_Name = "QRCodeTypePreparation"
Option Explicit
Private SettingName As String



Public Function PrintLabelPreparation(ByVal PreparationFileName As String, ByVal PreparationID As Long, ByVal RecipeCode As String, ByVal bCODLine As Boolean)
Dim MyQRCode As Qr01
Dim PREP_PATH As String
Dim HannaCode As String
Dim MaxHannaCode As String
Dim i As Integer
Dim rc As Boolean
Dim stringNote As String
Dim m_rc As Boolean
Dim RfpFileName As String
Dim SFGLot As String
On Error GoTo ERR_PRINT


    CloseSettingDataFile
    
    rc = False
    

    
     If MyOperatore.Name = "" Then

    If frmLogin.DoShow Then
            
    Else
        Exit Function
    End If

    End If

    
    
  
    If F_InputBox.DoShow("Enter Preparation Note", "QRCode : " & RecipeCode, , , , stringNote) Then MyQRCode.Note = stringNote
     
    With dbTabPreparation
        .filter = ""
        .filter = "ID='" & PreparationID & "'"
        If .EOF Then
        Else
            RfpFileName = IIf(IsNull(Trim(!RfpFileName)), "", Trim(!RfpFileName))
            
            MyQRCode.Code = IIf(IsNull(Trim(!HannaCode)), "", Trim(!HannaCode))
            MyQRCode.Exp = IIf(IsNull(Trim(!ExpDate)), "", Trim(!ExpDate))
            MyQRCode.LotPreparation = IIf(IsNull(Trim(!Lot)), "", Trim(!Lot))
            MyQRCode.Recipe = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
            MyQRCode.Date = FormatDataLAT(FormatDateTime(Now(), vbShortDate))
            MyQRCode.Time = FormatDateTime(Now(), vbShortTime)
            MyQRCode.Operator = MyOperatore.Name
            MyQRCode.Line = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
            MyQRCode.QC = "Waiting"
            
          
        End If
    
    End With
    
    MyQRCode.Lot = MyQRCode.LotPreparation
    
    If RfpFileName <> "" Then
    
       If FileExists(USER_TEMP_PATH & RfpFileName) Then
            PREP_PATH = USER_TEMP_PATH
        ElseIf FileExists(USER_DATA_PATH & RfpFileName) Then
            PREP_PATH = USER_DATA_PATH
        Else
            rc = False
            PopupMessage 2, "Canno't find data file...", , True, RfpFileName
            Exit Function
        End If
     
   
    End If

    MaxHannaCode = GetSettingData(RfpFileName, "HannaCodes", "HannaCodesCount", 0, PREP_PATH)

    
    If MaxHannaCode = 0 Then
        rc = DoPrintLabelPreparation(MyQRCode, False)
        m_rc = DoPrintLabelPreparation(MyQRCode, True)
    Else
     
     
        Dim bHide As Boolean
        Dim Um As String
        For i = 1 To MaxHannaCode
        
            MyQRCode.Code = GetSettingData(RfpFileName, "HannaCode" & i, "Code", "", PREP_PATH)
            Um = GetSettingData(RfpFileName, "HannaCode" & i, "Um", "", PREP_PATH)
            MyQRCode.Qty = GetSettingData(RfpFileName, "HannaCode" & i, "QtyToProduce", 0, PREP_PATH) '& " " & um
            
            If bCODLine Then
                MyQRCode.Lot = GetSettingData(RfpFileName, "HannaCode" & i, "LotNumber", MyQRCode.LotPreparation, PREP_PATH) '& " " & um
                If MyQRCode.Lot <> "" Then
                    SFGLot = IIf(IsNull(Trim(dbTabPreparation!SFGLot)), "", Trim(dbTabPreparation!SFGLot))
                    dbTabPreparation!SFGLot = SetListOfString(SFGLot, MyQRCode.Lot)
                    dbTabPreparation.Update
                End If
               MyQRCode.LotPreparation = MyQRCode.Lot
                 
            End If
            
            
            bHide = GetSettingData(RfpFileName, "HannaCode" & i, "bHide", True, PREP_PATH)
          
                If Not (bHide) Then ' se non è NASCOSTO
                    If MyQRCode.Qty <> "" Then ' ma sopratutto se la QTY è > 0
                        rc = DoPrintLabelPreparation(MyQRCode, False)
                        'If Not (m_rc) Then
                            If F_MsgBox.DoShow("Print Label for Chemical QC?", MyQRCode.Code) Then
                                m_rc = DoPrintLabelPreparation(MyQRCode, True)
                            End If
                        'End If
                    End If
                    
                    
                    
                End If
    
        
        Next
    End If
    
ERR_END:
        
    If rc Then
        PopupMessage 2, "Printing Label....", MyQRCode.Code
    End If
    CloseSettingDataFile
    Exit Function
ERR_PRINT:
    MsgBox err.Description
    Resume Next
    
    
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
      .Code = IIf(strCode = "", .Code, strCode)
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

