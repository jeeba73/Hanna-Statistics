Attribute VB_Name = "modPrinter"
Option Explicit
Public PRINTER_PORT_NAME As String
Public PRINTERNAME As String
Public bStampanteSelezionata As Boolean

Public Function SelezionoStampante() As Boolean
Dim sName As String

sName = Printer.DeviceName ''

If F_PRINTER_SETTING.DoShow Then
    If F_MsgBox.DoShow("Printer : " + vbCrLf + Printer.DeviceName + vbCrLf + "Set as default Printer?", "Label Printer") Then
        SaveSetting App.Title, "LABEL PRINTER", "NAME", Printer.DeviceName
        PRINTER_PORT_NAME = Printer.Port
        SaveSetting App.Title, "LABEL PRINTER", "bStampanteSelezionata", True
        SelezionoStampante = True
    End If
End If


End Function
Public Function SelezionoStampanteSettings() As Boolean
Dim sName As String



If F_LABELPRINTER.DoShow Then
End If



End Function



Private Sub SetPrinter(ByVal MyName As String)
  Dim prn As Printer
  For Each prn In Printers
    If InStr(UCase(prn.DeviceName), UCase(MyName)) Then
      Set Printer = prn
      Exit For
    End If
  Next
End Sub

Public Function SearchInfoLabelPrinter() As Boolean
Dim rc As Boolean
 '------------------------------------------------------------
    ' controllo la stampante Etichette
    '------------------------------------------------------------
    
    
    MyPathLabel_Brother = GetSetting(App.Title, "PATH", "TEMPLATE LABEL", App.Path & "\ETICHETTA LAT.lbx")
    
start:
 
    If bStampanteSelezionata Then
        '-------------------------------------------------------
        '           PRINTER
        '-------------------------------------------------------
        'If bStampaOk Then
            
            PRINTERNAME = GetSetting(App.Title, "LABEL PRINTER", "NAME", "")
            
            If PRINTERNAME <> "" Then
                SetPrinter (PRINTERNAME)
                
            Else
                PopupMessage 2, "Select Printer...."
                GoTo sel:
            End If
        'End If
    
        If GetSetting(App.Title, "LABEL PRINTER", "bUtilizzo", False) Then
            PRINTERNAME = GetSetting(App.Title, "LABEL PRINTER", "NAME", "")
            If InStr(UCase(PRINTERNAME), UCase("ZDesigner")) Then
                bStampaOk = True
                
            ElseIf InStr(UCase(PRINTERNAME), "BROTHER") Then
                bStampaOk = CheckPrinterEtichetteBrother
            End If
           
        Else
            bStampaOk = False
             
             rc = False
            
        End If
        
    Else
sel:
        If SelezionoStampante Then
            bStampanteSelezionata = True
            
            GoTo start
        End If
    
    End If
    
    rc = bStampaOk
    
    SearchInfoLabelPrinter = rc
End Function

