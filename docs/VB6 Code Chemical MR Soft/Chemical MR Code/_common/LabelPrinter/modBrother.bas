Attribute VB_Name = "modBrother"
Option Explicit

Public MyPathLabel_Brother As String
Public MyPrinterName As String
Private prtidx As Integer
Private sPrinterName As String

Public bStampaOk As Boolean
Public Function BrotherPrinterExist() As Boolean
    '------------------------------------------------
    ' Brother Printer
    '------------------------------------------------
    Dim ErrStr As String
    Dim DestStr As String
    If SetBrotherPrinter(ErrStr, DestStr) Then
        bStampaOk = True
    Else
        bStampaOk = False
    End If
    BrotherPrinterExist = bStampaOk
End Function

Public Function CheckPrinterEtichetteBrother() As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim objDoc As bpac.Document
    rc = True
    On Error GoTo ERR_PRINTER
 '   DoPrintLabel 1, 2, 3, 4
    
 
    
   rc = BrotherPrinterExist
    
    
    
   
    
    If InStr(PRINTERNAME, "Brother") Then
 
    Set objDoc = CreateObject("bpac.Document")
    
    Else
        rc = False
        UploadDownloadMessageCounter = 0
        PopupMessage 2, "Attenzione accertarsi di aver installato una Stampante Brother", , True, "Brother Printer"
      
    End If
    
    
    
ERR_END:
    On Error GoTo 0
    CheckPrinterEtichetteBrother = rc
    Exit Function
ERR_PRINTER:
    rc = False
    PopupMessage 2, err.Description, , , "Brother Label Printer"
    'Call InstallaDriver
    
    Resume Next
End Function


Public Function SetTemplateLabel(ByVal Frm As Form) As Boolean
Dim rc As Boolean



    Dim szFilename As String
    szFilename = DialogFile(Frm.hWnd, 1, "Open", App.path & "\ETICHETTA.lbx", "Template Etichetta" & Chr(0) & "*.lbx" & Chr(0) & "Tutti i files" & Chr(0) & "*.*", "", "lbx")
    If szFilename = "" Then Exit Function
    
    MyPathLabel_Brother = szFilename
    DoEvents

    SaveSetting App.Title, "PATH", "TEMPLATE LABEL", MyPathLabel_Brother

        
End Function



Public Function DoPrintLabel(ByVal sData As String, ByVal Matricola As String, ByVal Certificato As String, ByVal Interna As String) As Boolean
Dim i As Integer
Dim t As Integer
Dim rc As Boolean
Dim objDoc As bpac.Document
    
    On Error GoTo ERR_PRINTER
    rc = True
    

    Set objDoc = CreateObject("bpac.Document")
    
    
    If (objDoc.Open(MyPathLabel_Brother) <> False) Then
    
        objDoc.GetObject("tCert").Text = Certificato
        
        objDoc.GetObject("tMatricola").Text = Matricola
        If Not (IsNull(Interna) Or Interna = "") Then
            objDoc.GetObject("tInterna").Text = Interna
        End If
        objDoc.GetObject("tData").Text = FormatDataLAT(sData)
        objDoc.StartPrint "", bpoDefault
        objDoc.PrintOut 1, bpoDefault
        objDoc.EndPrint
        objDoc.Close
        
        
        
        Call DoSaveEtichetta(sData, Matricola, Certificato, Interna)
    Else
        PopupMessage 2, "Impossibile trovare/stampare il file etichetta in :" & vbCrLf & MyPathLabel_Brother
    End If
    
ERR_END:
    On Error GoTo 0
    DoPrintLabel = rc
    bStampaOk = rc
    Exit Function
ERR_PRINTER:
    rc = False
    MessageCenterInfoTime = 2000
    PopupMessage 2, err.Description, , True
    DoEvents
   Resume Next
    
End Function

Public Function DoSaveEtichetta(ByVal sData As String, ByVal Matricola As String, ByVal Certificato As String, ByVal Interna As String) As Boolean
Dim i As Integer
Dim t As Integer
Dim rc As Boolean
Dim NomeFile As String

Dim objDoc As bpac.Document
    
    On Error GoTo ERR_PRINTER
    rc = True
    
     
    NomeFile = FormatNomeFile(Certificato & "_matr." & Matricola & "_" & sData)

    Set objDoc = CreateObject("bpac.Document")
    
    
    If (objDoc.Open(MyPathLabel_Brother) <> False) Then
    
    'objDoc.Printer.Name
    
        objDoc.GetObject("tCert").Text = Certificato
        
        objDoc.GetObject("tMatricola").Text = Matricola
        If Not (IsNull(Interna) Or Interna = "") Then
            objDoc.GetObject("tInterna").Text = Interna
        End If
        objDoc.GetObject("tData").Text = FormatDateTimeLAT(sData)

    Debug.Print USER_QUALITY_PATH & PathEtichette & "\" & NomeFile & ".bmp"
      
       rc = objDoc.Export(bexBmp, USER_QUALITY_PATH & PathEtichette & "\" & NomeFile & ".bmp", 900)
       rc = objDoc.Export(bexLbx, USER_QUALITY_PATH & PathEtichette & "\" & NomeFile & ".lbx", 900)
    
    
    ' objDoc.EndPrint
        objDoc.Close
        
    Else
        PopupMessage 2, "Impossibile trovare/stampare il file etichetta in :" & vbCrLf & MyPathLabel_Brother
    End If
    
ERR_END:
    On Error GoTo 0
    DoSaveEtichetta = rc
    If rc Then PopupMessage 2, "Etichetta creata correttamente..."
    bStampaOk = rc
    Exit Function
ERR_PRINTER:
    rc = False
    MessageCenterInfoTime = 2000
    PopupMessage 2, "Errore stampante.... Per Stampare il l'etichetta è necessario collegare la stampante : Disabilito la stampa", , True
    DoEvents
    GoTo ERR_END
    
End Function


       

Public Function InstallaDriver()

If FileExists(App.path & "\Brother\bcciw31006.msi") Then
   
        ApriEseguibile App.path & "\Brother\bcciw31006.msi"
    Else
        
End If
    
End Function

Public Function SetBrotherPrinter(ByRef ErrStr As String, ByVal Destinazione As String) As Boolean
 Dim prt As Printer
 Dim util As bpac.Document
Dim rc As Boolean
On Error GoTo SET_ERR:
 
 
    rc = True
    
    Set util = CreateObject("bpac.Document")
    sPrinterName = PRINTERNAME ' Printer.DeviceName '"Brother QL"
    
    Rem -- Configure the PDF print job
    prtidx = PrinterIndex(sPrinterName)
    If prtidx < 0 Then
        err.Raise 1000, , "No printer was found by the name of '" & sPrinterName & "'."
        ErrStr = "Non è stata trovata la stampante : '" & sPrinterName & "'."
        rc = False
    End If
    
 
    Rem -- Set the current printer
    Set Printer = Printers(prtidx)
    

    
RESUME_ERR:
    On Error GoTo 0
    SetBrotherPrinter = rc
    If rc Then
        Dim w As New WshNetwork
        w.SetDefaultPrinter (sPrinterName)
       UploadDownloadMessageCounter = 0
        PopupMessage 2, "Stampante riconosciuta correttamente....", , , sPrinterName
        Set w = Nothing
    End If
    Exit Function
SET_ERR:
    rc = False
    If InStr(UCase(err.Description), UCase("Brother")) Then
       If F_MsgBox.DoShow(err.Description & vbCrLf & "Installare il Driver della stampante Brother ...", "Driver Stampante PDF", False, "Installa", "Esci") Then
            
            InstallaDriver
            DoEvents
       
        Else
          PopupMessage 2, "Sarà necessario Installare i driver (BioPDF) per poter creare correttamente i documenti PDF."
          SaveSetting App.Title, "LABEL PRINTER", "bUtilizzo", False
        End If
    Else
      '  PopupMessage 2, err.Description, , True, "Driver Stampante Brother"
    End If
    Resume RESUME_ERR
End Function
