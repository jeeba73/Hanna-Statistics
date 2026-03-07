Attribute VB_Name = "modPDF"
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const SETTINGS_PROGID = "biopdf.PDFSettings"
Const UTIL_PROGID = "biopdf.PDFUtil"

Private prtidx As Integer
Private sPrinterName As String
Private settings As Object
Private util As Object
Public DefaultPrinter As String



Private Sub GetDefaultPrinter()

If GetSetting(App.Title, "Printer", "DefaultPrinter", "") <> "" Then
    DefaultPrinter = GetSetting(App.Title, "Printer", "DefaultPrinter", "")
Else
    DefaultPrinter = Printer.DeviceName
End If

End Sub




Public Function SetUserDefaultPrinter()


        Dim w As New WshNetwork
        w.SetDefaultPrinter (DefaultPrinter)
        Set w = Nothing
 
    


End Function

Public Function PrinterExist() As Boolean
    '------------------------------------------------
    ' BioPFD Printer
    '------------------------------------------------
    Dim ErrStr As String
    Dim DestStr As String
    If SetPDFPrinter(ErrStr, DestStr) Then
        bStampanteOK = True
    Else
        bStampanteOK = False
    End If
    PrinterExist = bStampanteOK
End Function

Private Function PrinterIndex(ByVal PRINTERNAME As String) As Integer
    Dim i As Integer
    
    For i = 0 To Printers.Count - 1
        If LCase(Printers(i).DeviceName) Like LCase(PRINTERNAME) Then
            PrinterIndex = i
            
            Exit Function
        End If
    Next
    PrinterIndex = -1
End Function

Public Function SetPDFPrinter(ByRef ErrStr As String, ByVal Destinazione As String) As Boolean
 Dim prt As Printer
Dim rc As Boolean
Dim MyErr As String
On Error GoTo SET_ERR:

    GetDefaultPrinter

    rc = True
    Set util = CreateObject(UTIL_PROGID)
    sPrinterName = util.defaultprintername
    'MsgBox UTIL_PROGID
    Rem -- Configure the PDF print job
    Set settings = CreateObject(SETTINGS_PROGID)
    'MsgBox SETTINGS_PROGID
    settings.PRINTERNAME = sPrinterName
       
    Debug.Print Destinazione
    settings.SetValue "Output", Destinazione & ".pdf"
    settings.SetValue "ConfirmOverwrite", "no"
    settings.SetValue "ShowSaveAS", "never"
    settings.SetValue "ShowSettings", "never"
    settings.SetValue "ShowPDF", "never"
    settings.SetValue "RememberLastFileName", "no"
    settings.SetValue "RememberLastFolderName", "no"
    settings.WriteSettings True
    
    Rem -- Find the index of the printer
    prtidx = PrinterIndex(sPrinterName)
    'MsgBox sPrinterName
    'MsgBox prtidx
    If prtidx < 0 Then
        err.Raise 1000, , "No printer was found by the name of '" & sPrinterName & "'."
        ErrStr = "Non è stata trovata la stampante : '" & sPrinterName & "'."
        rc = False
    End If
    

    Rem -- Set the current printer
    Set Printer = Printers(prtidx)
RESUME_ERR:
    On Error GoTo 0
    SetPDFPrinter = rc
    
    If rc Then
        Dim w As New WshNetwork
        w.SetDefaultPrinter (sPrinterName)
        Set w = Nothing
    Else
        ' verrà utilizzata la stampante di sistema
    End If
    Exit Function
SET_ERR:
    rc = False
    MyErr = err.Description
    PopupMessage 2, "Attenzione stampante virtuale Bio-pdf non installata correttamente.", , True
    If InStr(UCase(MyErr), UCase("bioPDF")) Then
       If F_MsgBox.DoShow(err.Description & vbCrLf & "Installare il Driver della stampante virtuale Bio-PDF per generare Report e Certificati.", "Driver Stampante PDF") Then
            
            InstallaDriver
            DoEvents
       
        Else
            PopupMessage 2, "Sarà necessario Installare i driver (BioPDF) per poter creare i documenti PDF."
        End If
    Else
        PopupMessage 2, "Errore Stampante : " & MyErr, , True, "Driver Stampante PDF"
    End If
    Resume RESUME_ERR
End Function

Private Function InstallaDriver()

If FileExists(App.Path & "\PDFPrinter\Setup_bioPDFSetup_11_4_0_2674_PRO_EXP.exe") Then
   
        ApriEseguibile App.Path & "\PDFPrinter\Setup_bioPDFSetup_11_4_0_2674_PRO_EXP.exe"
    Else
        PopupMessage 2, "Attenzione impossibile trovare Setup_bioPDFSetup_11_4_0_2674_PRO_EXP.exe, Si consiglia di Reinstallare il programma.", , True
End If
    
End Function
