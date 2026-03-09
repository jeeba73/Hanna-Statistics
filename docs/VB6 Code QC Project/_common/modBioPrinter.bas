Attribute VB_Name = "modPDF"
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const SETTINGS_PROGID = "biopdf.PDFSettings"
Const UTIL_PROGID = "biopdf.PDFUtil"

Private prtidx As Integer
Private sPrinterName As String
Private settings As Object
Private util As Object
Public bStampanteOK As Boolean
Public PDFPrinterName As String

Public Function PrinterIndex(ByVal PRINTERNAME As String) As Integer
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

Dim rc As Boolean
On Error GoTo SET_ERR:

    rc = True
    Set util = CreateObject(UTIL_PROGID)
    sPrinterName = util.defaultprintername
    PDFPrinterName = sPrinterName
    Rem -- Configure the PDF print job
    Set settings = CreateObject(SETTINGS_PROGID)
    settings.PRINTERNAME = sPrinterName
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
    If prtidx < 0 Then
        Err.Raise 1000, , "No printer was found by the name of '" & sPrinterName & "'."
        ErrStr = "No printer was found by the name of '" & sPrinterName & "'."
        rc = False
    End If
    
    Rem -- Set the current printer
    Set Printer = Printers(prtidx)
RESUME_ERR:
    On Error GoTo 0
    SetPDFPrinter = rc
    Exit Function
SET_ERR:
    rc = False
    MsgBox Err.Description
    Resume Next
End Function

Public Function SetWindowsPDFPrinter(ByRef ErrStr As String) As Boolean

Dim rc As Boolean
On Error GoTo SET_ERR:

    rc = True
    Set util = CreateObject(UTIL_PROGID)
    sPrinterName = util.defaultprintername
    PDFPrinterName = sPrinterName
    Rem -- Configure the PDF print job
    Set settings = CreateObject(SETTINGS_PROGID)
    settings.PRINTERNAME = sPrinterName
    settings.SetValue "ConfirmOverwrite", "no"
    settings.SetValue "ShowSaveAS", "never"
    settings.SetValue "ShowSettings", "never"
    settings.SetValue "ShowPDF", "never"
    settings.SetValue "RememberLastFileName", "no"
    settings.SetValue "RememberLastFolderName", "no"
    settings.WriteSettings True
    
    Rem -- Find the index of the printer
    prtidx = PrinterIndex(sPrinterName)
    If prtidx < 0 Then
        Err.Raise 1000, , "No printer was found by the name of '" & sPrinterName & "'."
        ErrStr = "No printer was found by the name of '" & sPrinterName & "'."
        rc = False
    End If
    
    Rem -- Set the current printer
    Set Printer = Printers(prtidx)
RESUME_ERR:
    On Error GoTo 0
    SetWindowsPDFPrinter = rc
    Exit Function
SET_ERR:
    rc = False
    MsgBox Err.Description
    Resume Next
End Function


