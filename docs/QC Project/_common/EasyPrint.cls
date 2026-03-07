VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EasyPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********** EasyPrint ***************
'by Claudio Gucchierato
'e-mail: cgucchierato@libero.it
'Data: 05/2001
'Vers.: 1.0
'*********************************************************************
'Per una descrizione dettagliata sull'uso  dei Metodi e delle
'Proprietà della classe leggere il file EasyPrint.doc allegato.
'---------------------------------------------------------------------------------------

'N.B. --------------------------------------------------------------------------------
'Nota:  l 'effetto delle proprietà dell'oggetto Printer dipende dal
'driver fornito dalla casa produttrice  della stampante, quindi è
'possibile che alcune impostazioni non abbiano effetto oppure,
'che impostazioni al di fuori dell'intervallo riconosciuto, diano o
'non diano origine a un errore.
'-----------------------------------------------------------------------------------------
Option Explicit

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                         ByVal nIndex As Long) As Long
 
 'zona non stampabile
 Dim OffsetX  As Single
 Dim OffsetY  As Single

'Sono dichiarati solo i formati più usati. E' possibile personalizzare
'la lista aggiungendo qualunque altro formato.
Enum Dimensioni_Foglio
    A3 = vbPRPSA3                              ' A3, 297 x 420 mm
    A4 = vbPRPSA4                              ' A4, 210 x 297 mm
    A5 = vbPRPSA5                              ' A5, 148 x 210 mm
    B4 = vbPRPSB4                              ' B4, 250 x 354 mm
    B5 = vbPRPSB5                              ' B5, 182 x 257 mm
    Letter = vbPRPSLetter                      ' Letter 8.5 x 11 in.
    Tabloid = vbPRPSTabloid                  ' Tabloid, 11 x 17 in.
    Legal = vbPRPSLegal                       ' Legal, 8.5 x 14 in.
    Quarto = vbPRPSQuarto                   ' Quarto, 215 x 275 mm
    User = vbPRPSUser                         ' Definito dall 'utente.
End Enum
'Nota: Se si imposta la proprietà Height o Width per una stampante,
'PaperSize viene automaticamente impostata su vbPRPSUser.

Enum QualitàDiStampa
   Bozza = vbPRPQDraft                        '-1
   BassaRisoluzione = vbPRPQLow        '-2
   MediaRisoluzione = vbPRPQMedium   '-3
   AltaRisoluzione = vbPRPQHigh           '-4
End Enum
'Nota: su determinate stampanti, è possibile che alcune o tutte
'le impostazioni producano lo stesso risultato.

Enum TipoStampa
   Monocromatico = vbPRCMMonochrome  '1
   Colori = vbPRCMColor '2
End Enum

Enum OrientamentoDelFoglio
   Verticale = vbPRORPortrait           '1
   Orizzontale = vbPRORLandscape  '2
End Enum

Enum StileLinea
    Continua = vbSolid                           '(0)
    Lineetta = vbDash                            '(1)
    punto = vbDot                                  '(2)
    LineettaPunto = vbDashDot               '(3)
    LineettaPuntoPunto = vbDashDotDot '(4)
    'Invisibile = vbInvisible                        '(5)
    'Continua interna = vbInsideSolid        '(6)
End Enum
'Nota: se l'impostazione della proprietà DrawWidth è
'maggiore di 1, le impostazioni della proprietà DrawStyle
'da 1 a 4 produrranno sempre e solo una linea continua


Dim SpaziaturaLinee   As Single 'interlinea
Const m_DefaultSpaziatura = 1  ' mm
Const PI = 3.14159265 'pi greco

'******************  Proprietà della classe ****************************************
'----------------------------------------------------------------------------------------------------------

'Imposta l'interlinea (spaziatura fra una riga e la l'altra)
'(proprietà usata solo nelle stampe "multilinea", ovvero
'con stringhe che contengono almeno un'accapo (vbCr)
Property Let SpaziaturaLinea(Spaziatura As Single)
    If Spaziatura >= 0 Then
        SpaziaturaLinee = Spaziatura
    Else
        SpaziaturaLinee = m_DefaultSpaziatura '1mm
    End If
End Property
'Legge il valore dell'interlinea
Property Get SpaziaturaLinea() As Single
    SpaziaturaLinea = SpaziaturaLinee
End Property
' Imposta le dimensioni del font
Public Property Let FontSize(lFontSize As Integer)
    Printer.FontSize = lFontSize
End Property
'legge le dimensioni del font
Public Property Get FontSize() As Integer
    FontSize = Printer.FontSize
End Property
' Imposta il nome del font da usare
Public Property Let FontName(sFontName As String)
On Local Error GoTo FontError
    Printer.FontName = sFontName
   Exit Property
FontError:
End Property
' Legge il nome del font impostato
Public Property Get FontName() As String
    FontName = Printer.FontName
End Property
'Imposta la proprietà Bold per il font in uso
Public Property Let FontBold(bFontBold As Boolean)
    Printer.FontBold = bFontBold
End Property
'legge la proprietà Bold del font in uso
Public Property Get FontBold() As Boolean
    FontBold = Printer.FontBold
End Property
'Imposta la proprietà Italic per il font in uso
Public Property Let FontItalic(bFontItalic As Boolean)
    Printer.FontItalic = bFontItalic
End Property
'legge la proprietà Italic del font in uso
Public Property Get FontItalic() As Boolean
    FontItalic = Printer.FontItalic
End Property
'Imposta la proprietà Sottolinato per il font in uso
Public Property Let FontUnderline(bFontUnderline As Boolean)
    Printer.FontUnderline = bFontUnderline
End Property
'legge la proprietà Sottolinato del font in uso
Public Property Get FontUnderline() As Boolean
    FontUnderline = Printer.FontUnderline
End Property
'imposta il colore di primo piano
Public Property Let ForeColor(Color As Long)
 On Local Error GoTo ColorError
    Printer.ForeColor = Color
Exit Property
ColorError:
   Printer.ForeColor = vbBlack
End Property
'legge il il colore di primo piano
Public Property Get ForeColor() As Long
    ForeColor = Printer.ForeColor
End Property
'imposta le dimensioni del foglio
Public Property Let PaperSize(lPapersize As Dimensioni_Foglio)
    Printer.PaperSize = lPapersize
    'MsgBox Printer.DeviceName
End Property
'legge le dimensioni del foglio impostate
Property Get PaperSize() As Dimensioni_Foglio
    PaperSize = Printer.PaperSize
End Property
'imposta il numero di copie
Public Property Let Copies(lCopies As Integer)
    Printer.Copies = lCopies
End Property
'legge il numero di copie
Public Property Get Copies() As Integer
    Copies = Printer.Copies
End Property
'legge la qualità di stampa impostata
Public Property Let PrintQuality(Quality As QualitàDiStampa)
    Printer.PrintQuality = Quality
End Property
'imposta la modalità di stampa impostata (colore o B&W)
Public Property Let ColorMode(Output As TipoStampa)
    Printer.ColorMode = Output
End Property
'imposta l'orientamento del foglio impostato (verticale o orizzontale)
Public Property Let Orientation(X As OrientamentoDelFoglio)
    Printer.Orientation = X
End Property
'imposta la proprietà della linea (continua, punto, punto-linea etc.)
Public Property Let DrawStyle(Style As StileLinea)
    Printer.DrawStyle = Style
End Property
'imposta lo spessore della linea impostato
Public Property Let DrawWidth(Width As Integer)
   Printer.DrawWidth = Width
End Property
'Legge la posizione corrente di stampa X (orizzontale)
Property Get CurrentX() As Single
    CurrentX = Printer.CurrentX - OffsetX
End Property
'Imposta la posizione di stampa X (orizzontale)
Property Let CurrentX(sCurrentX As Single)
    Printer.CurrentX = sCurrentX + OffsetX
End Property
'Imposta la posizione di stampa Y (verticale)
Property Let CurrentY(sCurrentY As Single)
    Printer.CurrentY = sCurrentY + OffsetY
End Property
'Legge la posizione corrente di stampa Y (verticale)
Property Get CurrentY() As Single
    CurrentY = Printer.CurrentY - OffsetY
End Property
'Legge il numero di font disponibili per la stampante
Public Property Get FontCount() As Integer
    FontCount = Printer.FontCount
End Property
'legge, tra i font disponibili, il nome del font in base all'indice
Public Property Get Fonts(iFont As Integer) As String
    Fonts = Printer.Fonts(iFont)
End Property
'legge le dimensioni (disponibili per la stampa) verticali della pagina
Property Get PageHeight() As Single
    PageHeight = Printer.ScaleHeight - (OffsetY * 2)
End Property
'legge le dimensioni (disponibili per la stampa) orizzontali della pagina
Property Get PageWidth() As Single
    PageWidth = Printer.ScaleWidth - (OffsetX * 2)
End Property
'legge il numero di pagina corrente
Property Get Page() As Integer
    Page = Printer.Page
End Property

'***************** Metodi pubblici della classe **********************************
'---------------------------------------------------------------------------------

'PRINTLEFT ***********************************************************************
'Questa procedura esegue la stampa allineata a Sx del foglio
'Xpos = posizione orizzontale di inizio stampa (margine SX)
'Ypos = posizione verticale di inizio stampa
'Testo = stringa da stampare
'Rientro = eventuale rientro SX della prima riga(Opzionale)
'WidthMax =  larghezza massima del testo (Opzionale)
'NB se la stringa da stampare ha una lunghezza maggiore di
'WidthMax vengono eliminati i caratteri a DX
Public Sub PrintLeft(Xpos As Single, YPos As Single, Testo As String, _
                 Optional Rientro As Single = 0, Optional WidthMax As Integer = 0)
    On Local Error GoTo ErroreStampa
    Dim NumAccapo As Long
    Dim Inizio As Long
    Dim strTemp As String
    Dim TempXPos As Single
    
    NumAccapo = Accapo(Testo) 'verifico se sono presenti degli accapo (vbCr)
    
    TempXPos = Xpos + Rientro 'posizione di inizio stampa
    
    If NumAccapo = 0 Then 'se non sono presenti degli accapo
        If Xpos >= 0 Then Printer.CurrentX = OffsetX + TempXPos
        If YPos >= 0 Then Printer.CurrentY = OffsetY + YPos
        If WidthMax > 0 Then Testo = VerififyWidthTxt(Testo, WidthMax, True)
        Printer.Print Testo
    Else 'sono presenti degli accapo
        
        For Inizio = 0 To NumAccapo  'ricerca tutte le righe presenti
            If Inizio > 0 Then TempXPos = Xpos
            
            strTemp = GetLineFromString(Testo, Inizio)
             If WidthMax > 0 Then strTemp = VerififyWidthTxt(strTemp, WidthMax, True)
            If Xpos >= 0 Then Printer.CurrentX = OffsetX + TempXPos
            If YPos >= 0 Then Printer.CurrentY = (OffsetY + YPos) + (Inizio * Printer.TextHeight(strTemp) + Inizio * SpaziaturaLinee)
            Printer.Print strTemp
        Next Inizio
    End If
Exit Sub
ErroreStampa:
PrintError err.NUMBER
End Sub

'PRINTRIGHT ********************************************************************
'Questa procedura esegue la stampa allineata a DX del foglio
'Xpos = posizione orizzontale di inizio stampa (margine DX)
'Ypos = posizione verticale di inizio stampa
'Testo = stringa da stampare
'Rientro = eventuale rientro DX della prima riga(Opzionale)
'WidthMax =  larghezza massima del testo (Opzionale)
'NB se la stringa da stampare ha una lunghezza maggiore di
'WidthMax vengono eliminati i caratteri a SX
'se si vogliono eliminare i caratteri a DX, richiamare la funzione
'VerififyWidthTxt con l'ultimo argomento = True

Public Sub PrintRight(Xpos As Single, YPos As Single, Testo As String, _
                 Optional Rientro As Single = 0, Optional WidthMax As Integer = 0)
    On Local Error GoTo ErroreStampa
    Dim NumAccapo As Long
    Dim Inizio As Long
    Dim strTemp As String
    Dim TempXPos As Single
    
    NumAccapo = Accapo(Testo)
   TempXPos = Xpos - Rientro

    If NumAccapo = 0 Then
        If YPos >= 0 Then Printer.CurrentY = OffsetY + YPos
        If WidthMax > 0 Then Testo = VerififyWidthTxt(Testo, WidthMax, False)
        If Xpos >= 0 Then Printer.CurrentX = (OffsetX + TempXPos) - Printer.TextWidth(Testo)
        Printer.Print Testo
    Else
        
        For Inizio = 0 To NumAccapo
        
            If Inizio > 0 Then TempXPos = Xpos
            
            strTemp = GetLineFromString(Testo, Inizio)
            If WidthMax > 0 Then strTemp = VerififyWidthTxt(strTemp, WidthMax, False)
            If Xpos >= 0 Then Printer.CurrentX = (OffsetX + TempXPos) - Printer.TextWidth(strTemp)
            If YPos >= 0 Then Printer.CurrentY = (OffsetY + YPos) + (Inizio * Printer.TextHeight(strTemp) + Inizio * SpaziaturaLinee)
            Printer.Print strTemp
            
        Next Inizio
    End If
Exit Sub
ErroreStampa:
PrintError err.NUMBER

End Sub

'PRINTCENTRE ************************************************************
'Questa procedura esegue la stampa centrata rispetto al valore di Xpos
'Xpos = posizione orizzontale di inizio stampa (centro della stampa)
'Ypos = posizione verticale di inizio stampa
'Testo = stringa da stampare
'WidthMax =  larghezza massima del testo (Opzionale)
'NB se la sringa da stampare ha una lunghezza maggiore di
'WidthMax vengono eliminati i caratteri a DX
'se si vogliono eliminare i caratteri a SX, richiamare la funzione
'VerififyWidthTxt con l'ultimo argomento = False

Public Sub PrintCentre(Xpos As Single, YPos As Single, Testo As String, _
                 Optional WidthMax As Integer = 0)
    On Local Error GoTo ErroreStampa
    Dim NumAccapo As Long
    Dim lngStart As Long
    Dim strTemp As String
    Dim TempXPos As Single
    
    NumAccapo = Accapo(Testo)
    TempXPos = Xpos
    If NumAccapo = 0 Then
        If WidthMax > 0 Then Testo = VerififyWidthTxt(Testo, WidthMax, True)
        If Xpos >= 0 Then Printer.CurrentX = (OffsetX + TempXPos) - (Printer.TextWidth(Testo) / 2)
        If YPos >= 0 Then Printer.CurrentY = OffsetY + YPos
        Printer.Print Testo
    Else
        
        For lngStart = 0 To NumAccapo
            
            If lngStart > 0 Then TempXPos = Xpos
            
            strTemp = GetLineFromString(Testo, lngStart)
            If WidthMax > 0 Then strTemp = VerififyWidthTxt(strTemp, WidthMax, True)
        
            If Xpos >= 0 Then Printer.CurrentX = (OffsetX + TempXPos) - (Printer.TextWidth(strTemp) / 2)
            If YPos >= 0 Then Printer.CurrentY = (OffsetY + YPos) + (lngStart * Printer.TextHeight(strTemp) + lngStart * SpaziaturaLinee)
            Printer.Print strTemp
        Next lngStart
    End If
Exit Sub
ErroreStampa:
PrintError err.NUMBER
End Sub

'PRINTSHAPE *****************************************************************
'Questa procedura disegna una forma grafica tipo:
'rettangolo o quadrato, rettangolo o quadrato con angoli arrotondati, ovale o cerchio
'StartX = posizione orizzontale di inizio stampa (angolo sup SX)
'StartY = posizione verticale di inizio stampa(angolo superirore SX)
'ObjWidth = Larghezza della forma da disegnare
'ObjHeight = Altezza della forma da disegnare
'Rounded = smussatura degli angoli
'Impostando Rounded con un valore uguale o superiore al 50%
'del lato più piccolo si ottiene una forma circolare o ellittica
Public Sub PrintShape(StartX As Single, StartY As Single, ObjWidth As Single, _
                 ObjHeight As Single, Optional Rounded As Single = 0)

On Local Error GoTo ErroreStampa
    
    If Rounded > 0 Then
      Dim Temp As Single
        
        Temp = ObjWidth
        If ObjHeight <= ObjWidth Then Temp = ObjHeight
        If Rounded > Temp / 2 Then Rounded = Temp / 2
        
        Printer.CurrentX = StartX + OffsetX + Rounded
        Printer.CurrentY = StartY + OffsetY
        ' lato superiore
        Printer.Line Step(0, 0)-Step(ObjWidth - 2 * Rounded, 0)
        ' angolo sup. DX
        Printer.Circle Step(0, Rounded), Rounded, , 0, PI / 2
        ' lato DX
        Printer.Line Step(Rounded, 0)-Step(0, ObjHeight - 2 * Rounded)
        ' angolo inf. DX
        Printer.Circle Step(-Rounded, 0), Rounded, , 3 * PI / 2, 0
        ' lato inferiore
        Printer.Line Step(0, Rounded)-Step(-(ObjWidth - 2 * Rounded), 0)
        ' angolo inf. SX
        Printer.Circle Step(0, -Rounded), Rounded, , PI, 3 * PI / 2
        ' lato SX
        Printer.Line Step(-Rounded, 0)-Step(0, -(ObjHeight - 2 * Rounded))
        ' angolo sup. SX
        Printer.Circle Step(Rounded, 0), Rounded, , PI / 2, PI
    
    Else
        
        Printer.Line (StartX + OffsetX, StartY + OffsetY)-Step(ObjWidth, ObjHeight), , B
    End If
    
    Printer.CurrentX = StartX + OffsetX + ObjWidth
    Printer.CurrentY = StartY + OffsetY + ObjHeight
Exit Sub
ErroreStampa:
PrintError err.NUMBER
End Sub

'PRINTPICTURE ****************************************************************
'Questa procedura stampa un'immagine contenuta in un controllo Picture o Image
'StartX = posizione orizzontale di inizio stampa (lato sup SX)
'StartY = posizione verticale di inizio stampa (lato sup SX)
'objPicture = nome del controllo (es: Picture1 o Image1)
'Height = altezza (opzionale)
'Width = larghezza (opzionale)
'se Height e Width non vengono specificati l'immagine sarà stampata
'con le dimensioni originali

Public Sub PrintPicture(StartX As Single, StartY As Single, objPicture As Object, _
                 Optional Height As Single = 0, Optional Width As Single = 0)
    
   On Local Error GoTo ErroreStampa
    
    If Height > 0 And Width > 0 Then 'modificata altezza e/o larghezza
        
        Printer.PaintPicture objPicture.Image, StartX + OffsetX, _
                             StartY + OffsetY, Width, Height
    
    ElseIf Height = 0 And Width > 0 Then 'modifica solo larghezza
    
        Printer.PaintPicture objPicture.Image, StartX + OffsetX, _
                             StartY + OffsetY, Width
    
    ElseIf Height > 0 And Width = 0 Then 'modifica solo altezza
        
        Printer.PaintPicture objPicture.Image, StartX + OffsetX, _
                             StartY + OffsetY, , Height
    Else 'nessuna modifica, dimensioni originali
        
        Printer.PaintPicture objPicture.Image, StartX + OffsetX, _
                             StartY + OffsetY
    End If
Exit Sub
ErroreStampa:
PrintError err.NUMBER

End Sub
'Larghezza del testo in mm
Public Function TextWidth(sText As String) As Single
    TextWidth = Printer.TextWidth(sText)
End Function
'Altezza del testo in mm
Public Function TextHeight(sText As String) As Single
    TextHeight = Printer.TextHeight(sText)
End Function
'Stampa il documento
Public Sub EndDoc()
On Local Error GoTo err:
    Printer.EndDoc
err:
   ' MsgBox "Errore Stampante"
End Sub
'Inizia una nuova pagina
Public Sub NewPage()
On Local Error GoTo err:
   Printer.NewPage
err:
   ' MsgBox "Errore Stampante"

End Sub
'Termina l'operazione di stampa in corso.
Public Sub KillDoc()
   Printer.KillDoc
End Sub


'****************** Procedure private della classe *****************************
'--------------------------------------------------------------------------------------------------------

'Questa procedura verifica la lunghezza del testo
Private Function VerififyWidthTxt(txt As String, MaxWidth As Integer, l As Boolean) As String
If l = True Then 'elimina il testo dal lato DX
   Do While Printer.TextWidth(txt) > MaxWidth
      txt = Left(txt, Len(txt) - 1)
   Loop
Else 'elimina il testo dal lato SX
   Do While Printer.TextWidth(txt) > MaxWidth
      txt = Right(txt, Len(txt) - 1)
   Loop
End If
VerififyWidthTxt = txt
End Function

'Questa procedura estrae le singole linee di testo  presenti nella stringa
Private Function GetLineFromString(strText As String, NumLinea As Long) As String

    Dim Inizio As Long
    Dim strTemp As String
    Dim strLine As String
    
    strTemp = strText
    
    For Inizio = 0 To NumLinea
        strLine = ""
        
        'estrae la linea di testo carattere per carattere
        Do Until (Left(strTemp, 1) = vbCr) Or (Len(strTemp) = 0)
            strLine = strLine & Left(strTemp, 1)
            strTemp = Right(strTemp, Len(strTemp) - 1)
        Loop
        
        If Inizio = NumLinea Then
            GetLineFromString = Trim(strLine)
        Else
            If Len(strTemp) <> 0 Then
                'elimina il carattere accapo (carriage return)
                If Left(strTemp, 1) = vbCr Then strTemp = Right(strTemp, Len(strTemp) - 1)
            End If
            
            If Len(strTemp) <> 0 Then
                'elimina eventualmente il carattere line feed (vbLf)
                If Left(strTemp, 1) = vbLf Then strTemp = Right(strTemp, Len(strTemp) - 1)
            End If
            
            strTemp = Trim(strTemp)
        End If
        
    Next Inizio
    
End Function

'Questa procedura conta il numero di "accapo" (carriage returns) contenuti nel testo
Private Function Accapo(strText As String) As Long
  Dim Counter As Long
  Dim Pointer As Long
    
   Counter = 0
    Pointer = InStr(strText, vbCr) 'ricerca il carattere vbCr (accapo)
    
    Do Until Pointer = 0
        Counter = Counter + 1 'conta i caratteri accapo presenti
        Pointer = InStr(Pointer + 1, strText, vbCr) 'next
    Loop
    Accapo = Counter
End Function
'Procedura generica per la notifica di eventuali errori
Private Sub PrintError(Num As Integer)
   MsgBox "Errore nella routine di stampa" & vbCrLf & _
   "Errore numero: " & Num & vbCrLf & _
   Error(Num), vbCritical, "EasyPrint"
End Sub
Private Sub Class_Initialize()
    
    Printer.ScaleMode = vbMillimeters ' imposto "scalamode" in millimetri
    
    GetPrinterOffsets 'legge l'area non stampabile
   ' imposta la posizione iniziale di stampa
    Printer.CurrentX = OffsetX
    Printer.CurrentY = OffsetY

End Sub
'Area non stampabile
Private Sub GetPrinterOffsets()
 
Const PHYSICALOFFSETX = 112  ' Physical Printable Area x margin
Const PHYSICALOFFSETY = 113  '  Physical Printable Area y margin
Const LOGPIXELSX = 88        ' Pixel/pollice logico in X
Const LOGPIXELSY = 90        ' Pixel/pollice logico in Y

'millimetri/pollici
OffsetX = -25.4 * GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX) / _
                      GetDeviceCaps(Printer.hdc, LOGPIXELSX)

OffsetY = -25.4 * GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY) / _
                      GetDeviceCaps(Printer.hdc, LOGPIXELSY)

End Sub

