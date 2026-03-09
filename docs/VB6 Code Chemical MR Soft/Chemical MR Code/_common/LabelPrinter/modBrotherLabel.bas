Attribute VB_Name = "modBrotherLabel"
Option Explicit
Type Qr01

    Recipe  As String
    Code    As String
    MRType  As String
    Qty     As String
    Lot     As String
    Exp     As String
    Operator  As String
    Date    As String
    Time    As String
    QC      As String
    Note    As String
    Text3   As String
    ID      As String
    Line    As String
    STDn    As String
    STDv    As String
    Storage As String
    
    Bottle  As String
    SuppLot As String
    MNP     As String
    ExpMR   As String
    
   
    strAcquisitions  As String
    
    'information: ID, Prep. Date, Exp. Date, Value, MR Code, Bottle, Supp. Lot, MNP, Exp. MR.

End Type


Public QRCode01 As Qr01
Public QRCode01Clean As Qr01



Public Function DoPrintLabel(ByVal sCode As String, ByVal sLotto As String, ByVal sBottle As String, ByVal sData As String, Optional ByVal strEtichetta As String) As Boolean
Dim i As Integer
Dim t As Integer
Dim rc As Boolean
Dim objDoc As bpac.Document
    
    On Error GoTo ERR_PRINTER
    rc = True
    

        If strEtichetta = "" Then strEtichetta = "MRLabel"

        MyPathLabel_Brother = App.PATH & "\" & strEtichetta & ".LBX"

    
    Set objDoc = CreateObject("bpac.Document")
    
    
    
    If (objDoc.Open(MyPathLabel_Brother) <> False) Then
    
    
    
        Debug.Print objDoc.GetPrinterName
        objDoc.GetObject("QRCode").Text = sCode & sQRSeparator & sLotto & sQRSeparator & sBottle
        objDoc.GetObject("tText").Text = sCode & sQRSeparator & sLotto & sQRSeparator & sBottle
      '  objDoc.GetObject("tLotto").Text =
      '  objDoc.GetObject("tBottle").Text = sBottle
       'objDoc.GetObject("tData").Text = FormatDateTime(Now(), vbShortDate)

        objDoc.StartPrint "", bpoDefault
        objDoc.PrintOut 1, bpoDefault
        objDoc.EndPrint
        objDoc.Close
    End If
    
ERR_END:
    On Error GoTo 0
    DoPrintLabel = rc
    'bStampaOk = rc
    Exit Function
ERR_PRINTER:
    rc = False
    MessageInfoTime = 2000
    MsgBox Err.Description
    PopupMessage 2, "Errore stampante.... Per Stampare il l'etichetta è necessario collegare la stampante : Disabilito la stampa", , True
    DoEvents
    GoTo ERR_END
    
End Function
Public Function DoPrintQRCodeStockLabel(ByRef iWarehouseEntry As WareHouseEntry) As Boolean
Dim i As Integer
Dim t As Integer
Dim rc As Boolean
Dim strQR As String
Dim strText As String
Dim strText2 As String
Dim strEtichetta As String
Dim objDoc As bpac.Document
Dim Nomefile As String

    On Error GoTo ERR_PRINTER
    
    
    strEtichetta = "MRStockLabel"
    MyPathLabel_Brother = App.PATH & "\" & strEtichetta & ".LBX"
    
    With iWarehouseEntry
            rc = True
            
                    
                
        strQR = .MRCode & sQRSeparator & .MRValueConcentration & sQRSeparator & .Lot & sQRSeparator & .SupplierEXP & sQRSeparator & .U & sQRSeparator & .Bottle(0)
       ' strText = "Code :" & .MRCode & vbCrLf & "MR Value :" & .MRValueConcentration & vbCrLf & "MR Lot :" & .Lot & vbCrLf & "SupplierEXP(Date):" & .SupplierEXP & vbCrLf & "U :" & .U

        strText = "Code" & vbCrLf & .MRCode & vbCrLf & "MR Value" & vbCrLf & .MRValueConcentration & vbCrLf & "MR Lot" & vbCrLf & .Lot & vbCrLf & "SupplierEXP(Date)" & vbCrLf & .SupplierEXP & vbCrLf & "U :" & .U



        If .NumberBottle = 0 Then Exit Function
        For i = 1 To .NumberBottle
            If .NumberBottle = 1 Then
                strText2 = "# " & .Bottle(0)
            Else
                strText2 = "# " & i & " / " & .NumberBottle
            End If
                
                

            Set objDoc = CreateObject("bpac.Document")
            
            
            If (objDoc.Open(MyPathLabel_Brother) <> False) Then
            
                objDoc.GetObject("QrCode").Text = strQR
                objDoc.GetObject("tText").Text = strText
                objDoc.GetObject("tText2").Text = strText2


                objDoc.StartPrint "", bpoDefault
                
              
                
                objDoc.PrintOut 1, bpoDefault
                objDoc.EndPrint
                
                Nomefile = FormatNomeFile(.MRCode & "." & .Lot & "." & .SupplierEXP & "." & i)
                rc = objDoc.Export(bexBmp, USER_LABEL_PATH & Nomefile & ".bmp", 900)
                rc = objDoc.Export(bexLbx, USER_LABEL_PATH & Nomefile & ".lbx", 900)
                
                objDoc.Close
            End If
        Next
        
    End With
ERR_END:
    On Error GoTo 0
    DoPrintQRCodeStockLabel = rc
    bStampaOk = rc
    Exit Function
ERR_PRINTER:
    rc = False
    MsgBox Err.Description
    
    MessageInfoTime = 2000
    PopupMessage 2, "Errore stampante.... Per Stampare il l'etichetta è necessario collegare la stampante : Disabilito la stampa", , True
    DoEvents
    GoTo ERR_END
    
End Function

Public Function DoPrintQRCodeLabel(ByVal sCode As String, ByVal sLotto As String, ByVal sBottle As String, ByVal Exp As String, Optional ByVal strEtichetta As String) As Boolean
Dim i As Integer
Dim t As Integer
Dim rc As Boolean
Dim strQR As String
Dim objDoc As bpac.Document
    
    On Error GoTo ERR_PRINTER
    rc = True
    
    
        If strEtichetta = "" Then
            strEtichetta = "Bottle"
            strQR = sCode & IIf(sLotto <> "", sQRSeparator & sLotto, "") & IIf(sBottle <> "", sQRSeparator & sBottle, "")
            
        End If


        MyPathLabel_Brother = App.PATH & "\" & strEtichetta & ".LBX"
    

    
    Set objDoc = CreateObject("bpac.Document")
    
    
    If (objDoc.Open(MyPathLabel_Brother) <> False) Then
    
        objDoc.GetObject("QrCode").Text = strQR
        objDoc.GetObject("tCode").Text = sCode
        objDoc.GetObject("tLotto").Text = sLotto
        objDoc.GetObject("tBottle").Text = sBottle
        
        ' stampo la data MR_EXP ???? Da confermare!
        objDoc.GetObject("tData").Text = FormatDateTime(Exp, vbShortDate)


        objDoc.StartPrint "", bpoDefault
        objDoc.PrintOut 1, bpoDefault
        objDoc.EndPrint
        objDoc.Close
    End If
    
ERR_END:
    On Error GoTo 0
    DoPrintQRCodeLabel = rc
    bStampaOk = rc
    Exit Function
ERR_PRINTER:
    rc = False
    MessageInfoTime = 2000
    PopupMessage 2, "Errore stampante.... Per Stampare il l'etichetta è necessario collegare la stampante : Disabilito la stampa", , True
    DoEvents
    GoTo ERR_END
    
End Function





Public Function DoPrintStandardLabel(ByRef QRCode As Qr01, ByVal bToQC As Boolean, ByVal bManual As Boolean) As Boolean
Dim i As Integer
Dim t As Integer
Dim rc As Boolean
Dim QrString As String
Dim tText As String
Dim tText2 As String
Dim tNote As String
Dim Nomefile As String
Dim strEtichetta As String

Dim objDoc As bpac.Document
    
    On Error GoTo ERR_PRINTER
    rc = True
    

        If strEtichetta = "" Then strEtichetta = "Label01"

        MyPathLabel_Brother = App.PATH & "\" & strEtichetta & ".LBX"

    
    Set objDoc = CreateObject("bpac.Document")
    
   
    
    If (objDoc.Open(MyPathLabel_Brother) <> False) Then
    
         With QRCode
         
            .Text3 = "STD Production"
          
            If bManual Then
            
                tText = .Code & vbCrLf & _
                        .MRType & vbCrLf & _
                         "STDv: " & .STDv & vbCrLf & _
                         "Exp: " & .Exp & vbCrLf & _
                         "Storage: " & .Storage & vbCrLf & _
                         "Date: " & .Date & vbCrLf & _
                         "Oper.: " & .Operator
                         
                
                QrString = .ID & sQRSeparator & _
                        .Date & sQRSeparator & _
                        .Exp & sQRSeparator & _
                        .STDv & sQRSeparator & _
                        .strAcquisitions & sQRSeparator & _
                        .MRType
                       
            Else
                
                tText = .Code & vbCrLf & _
                         .MRType & vbCrLf & _
                         "STD: " & .STDn & vbCrLf & _
                         .MRType & vbCrLf & _
                         "STDv: " & .STDv & vbCrLf & _
                         "Exp: " & .Exp & vbCrLf & _
                         "Storage: " & .Storage & vbCrLf & _
                         "Date: " & .Date & vbCrLf & _
                         "Oper.: " & .Operator
                
                
                QrString = .Code & sQRSeparator & _
                        .STDn & sQRSeparator & _
                        .STDv & sQRSeparator & _
                        .ID & sQRSeparator & _
                        .Operator & sQRSeparator & _
                        .MRType
                       
            
            End If

            tText2 = "ID: " & .ID
    
            
                        
            tNote = "Note : " & .Note
             
            Nomefile = FormatNomeFile(.Text3 & "." & .Code & "." & .STDn & "." & .Time)
  
 
             objDoc.GetObject("QRCode").Text = QrString
             objDoc.GetObject("tText").Text = tText
             objDoc.GetObject("tText2").Text = tText2
             objDoc.GetObject("tNote").Text = tNote
             objDoc.GetObject("tText3").Text = .Text3
        
           '
             
            rc = objDoc.Export(bexBmp, USER_LABEL_PATH & Nomefile & ".bmp", 900)
           
            objDoc.StartPrint "", bpoDefault
            objDoc.PrintOut 1, bpoDefault
            objDoc.EndPrint
            objDoc.Close
         End With
    
    End If
    
ERR_END:
    On Error GoTo 0
    
    If rc Then PopupMessage 2, "Printing Label...." & IIf(bToQC, vbCrLf & "Label to QC!", ""), , , QRCode.Code
    
    DoPrintStandardLabel = rc
    'bStampaOk = rc
    Exit Function
ERR_PRINTER:
    rc = False
    MessageInfoTime = 2000
    MsgBox Err.Description
    PopupMessage 2, "Errore stampante.... Per Stampare il l'etichetta è necessario collegare la stampante : Disabilito la stampa", , True
    DoEvents
    GoTo ERR_END
    
End Function



