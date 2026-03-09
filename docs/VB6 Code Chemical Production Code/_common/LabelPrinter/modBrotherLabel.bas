Attribute VB_Name = "modBrotherLabel"
Option Explicit

Type Qr01

    Recipe      As String
    Code        As String
    TotalQty    As String
    Box         As String
    Qty         As String
    Lot         As String
    Exp         As String
    Operator    As String
    Date        As String
    Time        As String
    QC          As String
    Note        As String
    Text3       As String
    Line        As String
    
    LotPreparation  As String
    
End Type


Public QRCode01 As Qr01
Public QRCode01Clean As Qr01



Public Function DoPrintLabelPreparation(ByRef QRCode As Qr01, ByVal bToQC As Boolean) As Boolean
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

        MyPathLabel_Brother = App.Path & "\" & strEtichetta & ".LBX"

    
    Set objDoc = CreateObject("bpac.Document")
    
   
    
    If (objDoc.Open(MyPathLabel_Brother) <> False) Then
    
         With QRCode
         
            .Text3 = IIf(bToQC, "Preparation To QC - Valid", "Preparation")
            
            If bCODLine Then
                tText = "Recipe: " & .Recipe & vbCrLf & _
                        "Code: " & .Code & vbCrLf & _
                        "SFGLot: " & .Lot & vbCrLf & _
                        "Qty: " & .Qty & vbCrLf & _
                        "Exp: " & .Exp & vbCrLf & _
                        "Prep.Op.: " & .Operator
            Else
                tText = "Recipe: " & .Recipe & vbCrLf & _
                         "Code: " & .Code & vbCrLf & _
                         "Qty: " & .Qty & vbCrLf & _
                         "Lot: " & .LotPreparation & vbCrLf & _
                         "Exp: " & .Exp & vbCrLf & _
                         "Prep.Op.: " & .Operator
            
            End If
            
        
            tText2 = "Date Time: " & .Date & " - " & .Time & vbCrLf & _
                     "QC: " & .QC
    
            QrString = .Recipe & sQRSeparator & _
                        .Code & sQRSeparator & _
                        .Qty & sQRSeparator & _
                        .LotPreparation & sQRSeparator & _
                        .Exp & sQRSeparator & _
                        .Operator & sQRSeparator & _
                        .Date & sQRSeparator & _
                        Replace(.Time, ":", ".") & sQRSeparator & _
                        .QC
                        
             tNote = "Note : " & .Note
             
            Nomefile = FormatNomeFile(.Text3 & "." & .Code & "." & .Lot & "." & .Date & "." & .Time)
  
            Debug.Print QrString
             objDoc.GetObject("QRCode").Text = QrString
             objDoc.GetObject("tText").Text = tText
             objDoc.GetObject("tText2").Text = tText2
             objDoc.GetObject("tNote").Text = tNote
             objDoc.GetObject("tText3").Text = .Text3
             objDoc.GetObject("tLine").Text = .Line
           
             
            rc = objDoc.Export(bexBmp, USER_LABEL_PATH & Nomefile & ".bmp", 900)
           ' rc = objDoc.Export(bexLbx, USER_LABEL_PATH & Nomefile & ".lbx", 900)
       
             objDoc.StartPrint "", bpoDefault
             objDoc.PrintOut 1, bpoDefault
             objDoc.EndPrint
             objDoc.Close
         End With
    
    End If
    
ERR_END:
    On Error GoTo 0
    
    If rc Then PopupMessage 2, "Printing Label...." & IIf(bToQC, vbCrLf & "Label to QC!", ""), , , QRCode.Code
    
    DoPrintLabelPreparation = rc
    'bStampaOk = rc
    Exit Function
ERR_PRINTER:
    rc = False
    MessageInfoTime = 2000
    MsgBox err.Description
    PopupMessage 2, "Warning : Label Error, please check Label file.", , True
    DoEvents
    GoTo ERR_END
    
End Function


Public Function DoPrintLabelProduction(ByRef QRCode As Qr01) As Boolean
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
    

        If strEtichetta = "" Then strEtichetta = "Label02"

        MyPathLabel_Brother = App.Path & "\" & strEtichetta & ".LBX"

    
    Set objDoc = CreateObject("bpac.Document")
    
   
    
    If (objDoc.Open(MyPathLabel_Brother) <> False) Then
    
         With QRCode
         
            tText = "Recipe: " & .Recipe & vbCrLf & _
                     "Code: " & .Code & vbCrLf & _
                     "Lot: " & .Lot & vbCrLf & _
                     "Exp: " & .Exp & vbCrLf & _
                     "Prod.Op.: " & .Operator

            tText2 = "Date Time: " & .Date & " - " & .Time
            
            'QrString = .Recipe & sQRSeparator & _
                        .Code & sQRSeparator & _
                        .Lot & sQRSeparator & _
                        .Exp & sQRSeparator & _
                        .Operator & sQRSeparator & _
                        .Date & sQRSeparator & _
                        Replace(.Time, ":", ".") & sQRSeparator & _
                        .QC
            QrString = .Recipe & sQRSeparator & _
                        .Code & sQRSeparator & _
                        .Qty & sQRSeparator & _
                        .Lot & sQRSeparator & _
                        .Exp & sQRSeparator & _
                        .Operator & sQRSeparator & _
                        .Date & sQRSeparator & _
                        Replace(.Time, ":", ".") & sQRSeparator & _
                        .QC
                        
            
             tNote = "Note : " & .Note
             
            Nomefile = FormatNomeFile(.Text3 & "." & .Code & "." & .Lot & "." & .Date & "." & .Time)
  
 
             objDoc.GetObject("QRCode").Text = QrString
             objDoc.GetObject("tText").Text = tText
             objDoc.GetObject("tText2").Text = tText2
             objDoc.GetObject("tText3").Text = .Text3
             objDoc.GetObject("tNote").Text = tNote
             objDoc.GetObject("tLine").Text = .Line
           '
             
            rc = objDoc.Export(bexBmp, USER_LABEL_PATH & Nomefile & ".bmp", 900)
          '  rc = objDoc.Export(bexLbx, USER_LABEL_PATH & Nomefile & ".lbx", 900)
       
             objDoc.StartPrint "", bpoDefault
             objDoc.PrintOut 1, bpoDefault
             objDoc.EndPrint
             objDoc.Close
         End With
    
    End If
    
ERR_END:
    On Error GoTo 0
    If rc Then PopupMessage 2, "Printing Label....", , , QRCode.Code
    DoPrintLabelProduction = rc
    'bStampaOk = rc
    Exit Function
ERR_PRINTER:
    rc = False
    MessageInfoTime = 2000
    MsgBox err.Description
    PopupMessage 2, "Warning : Label Error, please check Label file.", , True
    DoEvents
    GoTo ERR_END
    
End Function


Public Function DoPrintLabelCloseProduction(ByRef QRCode As Qr01, Optional bValue As Boolean) As Boolean
Dim i As Integer
Dim t As Integer
Dim rc As Boolean
Dim QrString As String
Dim tText As String
Dim tText2 As String
Dim tNote As String
Dim Nomefile As String
Dim strEtichetta As String
Dim sString As String
Dim tString As String

Dim objDoc As bpac.Document
    
    On Error GoTo ERR_PRINTER
    rc = True
    

        If strEtichetta = "" Then strEtichetta = "Label03_Prod"

        MyPathLabel_Brother = App.Path & "\" & strEtichetta & ".LBX"

    
    Set objDoc = CreateObject("bpac.Document")
    
   
    
    If (objDoc.Open(MyPathLabel_Brother) <> False) Then
    
         With QRCode
            If bValue Then
                Nomefile = FormatNomeFile("PROD_TOTALQTY" & "." & .Code & "." & .Lot & "." & .Date & "." & .Time)
            Else
                 Nomefile = FormatNomeFile("PROD" & "." & .Code & "." & .Lot & ".BoxN." & .Box & "." & .Date & "." & .Time)
            End If
 
            QrString = .Code & sQRSeparator & .Lot & sQRSeparator & IIf(bValue, .TotalQty, .Qty) & sQRSeparator & .Exp & "~"
             objDoc.GetObject("QRCode").Text = QrString
             
             objDoc.GetObject("tCode").Text = .Code
             objDoc.GetObject("tQty").Text = IIf(bValue, .TotalQty, .Qty)
             objDoc.GetObject("tLot").Text = .Lot
             objDoc.GetObject("tExp").Text = .Exp
             objDoc.GetObject("tDate").Text = .Date
             objDoc.GetObject("tOperator").Text = .Operator
             
             
            objDoc.GetObject("tBox").Text = IIf(bValue, "", .Box)
            
            
            rc = objDoc.Export(bexBmp, USER_LABEL_PATH & Nomefile & ".bmp", 900)
          
              objDoc.StartPrint "", bpoDefault
             objDoc.PrintOut 1, bpoDefault
             objDoc.EndPrint
             objDoc.Close
         End With
    
    End If
    
ERR_END:
    On Error GoTo 0
    If rc Then
        If bValue Then
            PopupMessage 2, "Printing Label...." & vbCrLf & "TOTAL Quantity : " & QRCode.TotalQty, , , "Closed Production Label"
        Else
            PopupMessage 2, "Printing Label...." & vbCrLf & "BOX Quantity : " & QRCode.Qty, , , "Box n." & QRCode.Box
        End If
    End If
    DoPrintLabelCloseProduction = rc
    'bStampaOk = rc
    Exit Function
ERR_PRINTER:
    rc = False
    MessageInfoTime = 2000
    MsgBox err.Description
    PopupMessage 2, "Warning : Label Error, please check Label file.", , True
    DoEvents
    Resume Next
  '  GoTo ERR_END
    
End Function
