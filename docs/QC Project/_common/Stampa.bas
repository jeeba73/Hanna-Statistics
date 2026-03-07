Attribute VB_Name = "Stampa"
Option Explicit

Dim Pr As New EasyPrint 'istanzio la classe
Dim PrCartel As New EasyPrint
Dim Dx As Single 'utilizzata come valore del margine DX
Dim Sx As Single 'utilizzata come valore del margine SX
Dim AltCar As Single ' altezza del carattere
Public check_temp As Integer
Private NumIspezioni As Integer
Private bUserValue As Boolean
Private devst(100) As Double
Private MeasurementUnit As String

    

Public Sub ImpostaStampa(Optional bValue As Boolean = True, Optional bCpk As Boolean = False)
    
    bUserValue = bValue

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub



Private Sub FormatFontDettaglio()
   
   'formatto il carattere per la stampa dei dati
   With Pr
    .FontBold = False
    .FontItalic = False
    .FontSize = 10
    AltCar = .TextWidth("ab")   'leggo l'altezza (in mm) del font ustato
   End With

End Sub

Private Sub StampaModulo(ByVal sString As String)  'stampa la struttura del modulo

    With Pr
       'stampo il logo (13x13mm) a 15 mm left e 15 mm top
       '.PrintPicture Sx, 15, imgLogo
       
       'personalizzo le caratteristiche del font
             .FontName = "Tahoma"
             .FontBold = True
             .FontUnderline = False
             .FontSize = 16
            ' .PrintLeft 20, 10, Societa
             .FontBold = True
             .FontSize = 16
             .PrintLeft 20, 20, ("HANNA INSTRUMENTS srl - Chemical QC")
            ' .FontSize = 14
            ' .PrintLeft Sx - 20, 20, Date
       'stampo parte dell'intestazione con gli attributi impostati
       '.PrintLeft 20, 20, Societa
      
       'stampo una doppia linea continua sotto all'intestazione
       .DrawStyle = Continua 'stile della linea
       .DrawWidth = 1 'spessore della linea
       .PrintShape Sx, 30, Dx, 0 'stampo la prima linea
       .PrintShape Sx, 31, Dx, 0 'stampo la seconda linea  1 mm più sotto
    
    
       'disegno il divisorio orizzontale superiore
       .DrawWidth = 2 'spessore della linea
       .PrintShape Sx, 78, Dx, 0 'disegna la linea
       
       .FontSize = 14
       .PrintCentre (Dx + Sx) / 2, 69, MyChemicalQC.HannaCode.code & " - " & sString
       
       .PrintShape Sx, 66, Dx, 0 'disegna la linea
       'disegno il divisorio orizzontale inferiore
      
    
    End With
End Sub

Public Sub DefaultValue(ByVal GrString As String, ByVal Index As Integer)
On Local Error GoTo err:

    With Pr
    'impostazioni generiche
       .PrintQuality = AltaRisoluzione
       .ColorMode = vbPRCMColor
       .PaperSize = A4
       .Orientation = Orizzontale
       Sx = 15 'bordo SX a 15 mm dal foglio
       Dx = Pr.PageWidth - 15 - Sx 'bordo dx a 15 mm dal foglio
    End With
        
error_label:

Exit Sub
err:
MsgBox ("Errore stampante.")
Resume error_label
End Sub

Public Sub Intestazione()
Dim passo
With Pr
     
   .FontBold = True
   .FontSize = 14
      
      passo = -5

   passo = passo + .TextWidth("ab")
    
   .FontBold = False
   .FontSize = 9
   .PrintLeft 15, 35 + passo, "Code : " & MyChemicalQC.HannaCode.code, 1
   passo = passo + AltCar + 1
   .PrintLeft 15, 35 + passo, "Description : " & MyChemicalQC.HannaCode.Description, 1
   passo = passo + AltCar + 1
   .PrintLeft 15, 35 + passo, "Lot : " & MyChemicalQC.Lot, 1
   passo = passo + AltCar + 1
   .FontSize = 9
   .PrintLeft 15, 35 + passo, "Recipe : " & MyChemicalQC.HannaCode.Recipe, 1
   passo = passo + AltCar
    .PrintLeft 15, 35 + passo, ("Exp : ") & MyChemicalQC.Exp, 1
   passo = passo + AltCar
    .PrintLeft 15, 35 + passo, ("Prod - First Day : ") & MyChemicalQC.ProdFirst, 1
    .PrintLeft 60, 35 + passo, ("Prod - Last Day : ") & MyChemicalQC.ProdLast, 1
   passo = passo + AltCar
 
 End With
End Sub


Public Sub PrintGrafico(ByVal Grafico As PictureBox, str_grafico As String, Optional Index As Integer)
Dim UserDecimal As String
Dim passo As Single 'distanza tra le righe

   'stampo la struttura del modulo

    
    Call FormatFontDettaglio
    Call Intestazione
    
    
    Call StampaModulo(str_grafico)
    'Call TabellaIspezioni
    
    'formatto i caratteri per la stampa dei dati
    
    MeasurementUnit = MyGraphicCheck.HannaCode.MeasurementUnit
    UserDecimal = MyGraphicCheck.HannaCode.Decimal
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
With Pr
    .FontBold = False
    .FontUnderline = False
    .FontSize = 9
    '.PrintCentre Dx / 2, 70, str_grafico & "  ( Machine : " & MyChemicalQC.MACHINE.Name & " - H" & Index & " )"
    .PrintPicture 3, 85, Grafico
    
    passo = passo + AltCar
    .PrintLeft 15, 170 + passo, ("STD Number : ") & MyGraphicCheck.DataControllo(Index).STDNumber, 1
    .PrintLeft 70, 170 + passo, ("# Selected Readings : ") & MyGraphicCheck.STDtest(Index).SelReadings, 1
    .PrintLeft 130, 170 + passo, ("# All Readings : ") & MyGraphicCheck.STDtest(Index).MaxReadings, 1
    
    passo = passo + AltCar + 1
    .PrintLeft 15, 170 + passo, ("STD Value : ") & Format$(MyGraphicCheck.DataControllo(Index).STDRef, UserDecimal) & " " & MeasurementUnit, 1
    .PrintLeft 70, 170 + passo, ("# Selected Test : ") & MyGraphicCheck.STDtest(Index).SelTest, 1
    .PrintLeft 130, 170 + passo, ("# All Test : ") & MyGraphicCheck.STDtest(Index).NumTest, 1
    
    
    
    passo = passo + AltCar + 1
    .PrintLeft 15, 170 + passo, ("Min Value : ") & MyGraphicCheck.DataControllo(Index).STDMin & " " & MeasurementUnit, 1
    .PrintLeft 70, 170 + passo, ("# Out Of Range : ") & MyGraphicCheck.DataControllo(Index).OutOfRangeData & " ( " & MyGraphicCheck.DataControllo(Index).OutOfRangeDataPerc & "% )", 1
    passo = passo + AltCar + 1
    .PrintLeft 15, 170 + passo, ("Max Value : ") & MyGraphicCheck.DataControllo(Index).STDMax & " " & MeasurementUnit, 1
    .PrintLeft 70, 170 + passo, ("Mean : ") & Format$(MyGraphicCheck.DataControllo(Index).media, UserDecimal) & " " & MeasurementUnit, 1
    .PrintRight Dx + Sx, 170 + passo, ("Operator : ") & MyGraphicCheck.DataControllo(Index).Operator
    passo = passo + AltCar
    
    .EndDoc 'invio il comando di stampa

End With

Set Pr = Nothing
End Sub
