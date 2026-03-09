Attribute VB_Name = "modBrotherLabel"
Option Explicit

Type Qr01

    Recipe  As String
    Code    As String
    Qty     As String
    Lot     As String
    Exp     As String
    Operator  As String
    date    As String
    Time    As String
    QC      As String
    Note    As String
    Text3   As String
    Tablet  As String
    Line    As String
    
End Type


Public QRCode01 As Qr01
Public QRCode01Clean As Qr01



Public Function DoPrintLabelReadings(ByRef QRCode As Qr01, ByVal bToQC As Boolean) As Boolean
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
         
            .Tablet = MyWorkStation.Workstation
            
            .Text3 = "QC Validation"
         
            tText = "Recipe: " & .Recipe & vbCrLf & _
                     "Code: " & .Code & vbCrLf & _
                     "Lot: " & .Lot & vbCrLf & _
                     "Exp: " & .Exp & vbCrLf & _
                     "Tablet: " & .Tablet & vbCrLf & _
                     "QC.Op.: " & .Operator

            tText2 = "Date Time: " & .date & " - " & .Time & vbCrLf & _
                     "QC: " & .QC
   
      
            QrString = .Recipe & sQRSeparator & _
                        .Code & sQRSeparator & _
                        .Note & sQRSeparator & _
                        .Lot & sQRSeparator & _
                        .Exp & sQRSeparator & _
                        .Operator & sQRSeparator & _
                        .date & sQRSeparator & _
                        Replace(.Time, ":", ".") & sQRSeparator & _
                         .QC & sQRSeparator & _
                          .Tablet
                        
             tNote = "Note : " & .Note
             
            Nomefile = FormatNomeFile(.Text3 & "." & .Code & "." & .Lot & "." & .date & "." & .Time)
  
 
             objDoc.GetObject("QRCode").Text = QrString
             objDoc.GetObject("tText").Text = tText
             objDoc.GetObject("tText2").Text = tText2
             objDoc.GetObject("tNote").Text = tNote
             objDoc.GetObject("tText3").Text = .Text3
             objDoc.GetObject("tLine").Text = .Line
    
    
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
    
    DoPrintLabelReadings = rc
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


