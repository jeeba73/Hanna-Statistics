Attribute VB_Name = "Stampa"
Option Explicit

Dim Pr As New EasyPrint 'istanzio la classe
Dim PrCartel As New EasyPrint
Dim dx As Single 'utilizzata come valore del margine DX
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
             .PrintLeft 20, 20, ("HANNA INSTRUMENTS srl - Chemical STDPreparation")
            ' .FontSize = 14
            ' .PrintLeft Sx - 20, 20, Date
       'stampo parte dell'intestazione con gli attributi impostati
       '.PrintLeft 20, 20, Societa
      
       'stampo una doppia linea continua sotto all'intestazione
       .DrawStyle = Continua 'stile della linea
       .DrawWidth = 1 'spessore della linea
       .PrintShape Sx, 30, dx, 0 'stampo la prima linea
       .PrintShape Sx, 31, dx, 0 'stampo la seconda linea  1 mm più sotto
    
    
       'disegno il divisorio orizzontale superiore
       .DrawWidth = 2 'spessore della linea
       .PrintShape Sx, 78, dx, 0 'disegna la linea
       
       .FontSize = 14
     '  .PrintCentre (dx + Sx) / 2, 69, STDPreparation.HannaCode.Code & " - " & sString
       
       .PrintShape Sx, 66, dx, 0 'disegna la linea
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
       dx = Pr.PageWidth - 15 - Sx 'bordo dx a 15 mm dal foglio
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
  ' .PrintLeft 15, 35 + passo, "Code : " & STDPreparation.HannaCode.Code, 1
   passo = passo + AltCar + 1
  ' .PrintLeft 15, 35 + passo, "Description : " & STDPreparation.HannaCode.Description, 1
   passo = passo + AltCar + 1
  ' .PrintLeft 15, 35 + passo, "Lot : " & STDPreparation.Lot, 1
   passo = passo + AltCar + 1
   .FontSize = 9
  ' .PrintLeft 15, 35 + passo, "Recipe : " & STDPreparation.HannaCode.Recipe, 1
   passo = passo + AltCar
  '  .PrintLeft 15, 35 + passo, ("Exp : ") & STDPreparation.Exp, 1
   passo = passo + AltCar
   ' .PrintLeft 15, 35 + passo, ("Prod - First Day : ") & STDPreparation.ProdFirst, 1
   ' .PrintLeft 60, 35 + passo, ("Prod - Last Day : ") & STDPreparation.ProdLast, 1
   passo = passo + AltCar
 
 End With
End Sub
