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
    
    
    
   
    
    If InStr(Printer.DeviceName, "Brother") Then
 
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
    PopupMessage 2, Err.Description, , , "Brother Label Printer"
    'Call InstallaDriver
    
    Resume Next
End Function


Public Function SetTemplateLabel(ByVal Frm As Form) As Boolean
Dim rc As Boolean



    Dim szFilename As String
    szFilename = DialogFile(Frm.hwnd, 1, "Open", App.PATH & "\ETICHETTA.lbx", "Template Etichetta" & Chr(0) & "*.lbx" & Chr(0) & "Tutti i files" & Chr(0) & "*.*", "", "lbx")
    If szFilename = "" Then Exit Function
    
    MyPathLabel_Brother = szFilename
    DoEvents

    SaveSetting App.Title, "PATH", "TEMPLATE LABEL", MyPathLabel_Brother

        
End Function



Public Function InstallaDriver()

If FileExists(App.PATH & "\Brother\bcciw31006.msi") Then
   
        ApriEseguibile App.PATH & "\Brother\bcciw31006.msi"
    Else
        
End If
    
End Function

Public Function SetBrotherPrinter(ByRef ErrStr As String, ByVal Destinazione As String) As Boolean
 Dim prt As Printer
 Dim util As bpac.Document
Dim rc As Boolean
On Error GoTo SET_ERR:
   ' MsgBox Printer.DeviceName
    rc = True
    Set util = CreateObject("bpac.Document")
    sPrinterName = Printer.DeviceName '"Brother QL"
    
    Rem -- Configure the PDF print job
    prtidx = PrinterIndex(sPrinterName)
    If prtidx < 0 Then
        Err.Raise 1000, , "No printer was found by the name of '" & sPrinterName & "'."
        ErrStr = "Non è stata trovata la stampante : '" & sPrinterName & "'."
        rc = False
    End If
    
    Debug.Print
    

    Rem -- Set the current printer
    Set Printer = Printers(prtidx)
    

    
RESUME_ERR:
    On Error GoTo 0
    SetBrotherPrinter = rc
    If rc Then
        Dim w As New WshNetwork
        w.SetDefaultPrinter (sPrinterName)
        PopupMessage 2, "Printer Settings OK....", , , sPrinterName
        'PopupMessage 2, "Imposto la stampante " & sPrinterName & " come predefinita", , , "Etichette/Label"
        Set w = Nothing
    End If
    Exit Function
SET_ERR:
    rc = False
    If InStr(UCase(Err.Description), UCase("Brother")) Then
       If F_MsgBox.DoShow(Err.Description & vbCrLf & "Installare il Driver della stampante Brother ...", "Driver Stampante PDF", False, "Installa", "Esci") Then
            
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
