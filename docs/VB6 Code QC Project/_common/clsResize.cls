VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ======================================================================================== '
' Componente: clsResize                                                 Creata: 18/11/2002 '
' Nome del File: clsResize.cls                                Autore: Massimiliano Mazzoli '
'                                                                                          '
' Scopo: Ridimensionare e Riposizionare tutti i controlli all'interno di una form          '
' ======================================================================================== '
' Commenti: la classe è stata creata prendendo spunto da altre due trovate in internet.    '                                                               '
' ======================================================================================== '

Option Explicit
' ======================================================================================== '
' Dichiarazione API.                                                                        '
' ======================================================================================== '
' Questa API previene dall'aggiornamento della Form in determinate situazione
Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long

' ======================================================================================== '
' Dichiarazioni Costanti.                                                                   '
' ======================================================================================== '
Private Const ssTabOffset = -75000              ' Offeset per i controlli ssTab

' ======================================================================================== '
' Dichiarazioni Tipi.                                                                       '
' ======================================================================================== '
Private Type udtControlloDati
    lngLeft As Long                     ' Valore iniziale della proprietà Left del controllo
    lngTop As Long                      ' Valore iniziale della proprietà Top del controllo
    lngWidth As Long                    ' Valore iniziale della proprietà Width del controllo
    lngHeight As Long                   ' Valore iniziale della proprietà Height del controllo
    lngIndex As Long                    ' Contiene il numero progressivo del controllo all'interno della form
    lngLevel As Long                    ' Contiene il livello del controllo all'interno di un container
    lngGrigliaWidth As Integer          ' Contiene l'ampiezza della griglia qual'ora lo sia
    lngGrigliaTag As String             ' Contiene il valore contenuto nel campo Tag
    lngContainerWidth As Long           ' Contiene la larghezza del container che contiene l'oggetto
    lngContainerHeight As Long          ' Contiene l'altezza del container che contiene l'oggetto
End Type

' ======================================================================================== '
' Dichiatazioni Enumerazioni.                                                              '
' ======================================================================================== '

' ======================================================================================== '
' Dichiarazioni Eventi.                                                                    '
' ======================================================================================== '

' ======================================================================================== '
' Dichiarazioni Variabili Pubbliche.                                                       '
' ======================================================================================== '

' ======================================================================================== '
' Dichiarazioni Variabili Private.                                                         '
' ======================================================================================== '
Private ctrControlli() As udtControlloDati      ' Contiene tutti i controlli della Form
Private lngFormWidthMin As Long                 ' Contiene la larghezza minima della Form oltre il quale non si può ridimensionare
Private lngFormHeightMin As Long                ' Contiene l'altezza minima della Form oltre il quale non si può ridimensionare

' ========================================================================================= '
' Routine: Class_Initialize                                              Creata: 18/11/2002 '
' Visibilità: Private                                        Autore: Massimiliano Mazzoli   '
' Descrizione: Costruttore utilizzato quando viene creata un'istanza della classe,          '
'              cioè un oggetto.                                                             '
' ========================================================================================= '
Private Sub Class_Initialize()
    Erase ctrControlli
    lngFormHeightMin = 0
    lngFormWidthMin = 0
End Sub

' ======================================================================================== '
' Routine: Class_Terminate                                              Creata: 18/11/2002 '
' Visibilità: Private                                         Autore: Massimiliano Mazzoli '
' Descrizione: Distruttore usato quando un'istanza di questa classe è distrutta.           '
' ======================================================================================== '
Private Sub Class_Terminate()
    Erase ctrControlli
    lngFormHeightMin = 0
    lngFormWidthMin = 0
End Sub

' ========================================================================================= '
' Routine: Inizializza                                                   Creata: 18/11/2002 '
' Visibilità: Public                                         Autore: Massimiliano Mazzoli   '
' Descrizione: Memorizza le dimensioni e le posizioni di tutti i controlli inseriti nella   '
'              Form.                                                             '
' ========================================================================================= '
Public Sub Inizializza(ByRef frmForm As Form, _
                       Optional lngHeightMin As Long = 0, _
                       Optional lngWidthMin As Long = 0)
    
    Dim lngCont01 As Long
    Dim lngLeft As Long
    Dim lngCont02 As Long
    Dim ctrControlloTemp As udtControlloDati
    Dim strTag As String
    
    If (lngHeightMin <> 0) Then
        lngFormHeightMin = lngHeightMin
    End If
    
    If (lngWidthMin <> 0) Then
        lngFormWidthMin = lngWidthMin
    End If
        
    Erase ctrControlli
    
    ReDim ctrControlli(0 To (frmForm.Controls.Count - 1))
    
    For lngCont01 = 0 To frmForm.Controls.Count - 1
        
        ctrControlli(lngCont01).lngIndex = lngCont01
        ctrControlli(lngCont01).lngGrigliaWidth = -1
        ctrControlli(lngCont01).lngGrigliaTag = ""
        ctrControlli(lngCont01).lngLevel = LivelloConta(frmForm.Controls(lngCont01))
        ctrControlli(lngCont01).lngContainerHeight = frmForm.Controls(lngCont01).Container.Height
        ctrControlli(lngCont01).lngContainerWidth = frmForm.Controls(lngCont01).Container.Width
        
        If (TypeOf frmForm.Controls(lngCont01) Is Line) Then
            ctrControlli(lngCont01).lngLeft = frmForm.Controls(lngCont01).X1
            ctrControlli(lngCont01).lngTop = frmForm.Controls(lngCont01).Y1
            ctrControlli(lngCont01).lngWidth = frmForm.Controls(lngCont01).X2
            ctrControlli(lngCont01).lngHeight = frmForm.Controls(lngCont01).Y2
        Else
          '  On Error Resume Next
            lngLeft = frmForm.Controls(lngCont01).Left
            
            If (lngLeft < 0) Then
                lngLeft = Abs(ssTabOffset - frmForm.Controls(lngCont01).Left)
            End If

            ctrControlli(lngCont01).lngLeft = lngLeft
            ctrControlli(lngCont01).lngTop = frmForm.Controls(lngCont01).TOp
            ctrControlli(lngCont01).lngWidth = frmForm.Controls(lngCont01).Width
            ctrControlli(lngCont01).lngHeight = frmForm.Controls(lngCont01).Height

            strTag = TagEstrai(frmForm.Controls(lngCont01).Tag, "rs=")
            strTag = UCase(strTag)
            
            If (CBool(InStr(strTag, "G"))) Then
             
                ctrControlli(lngCont01).lngGrigliaWidth = frmForm.Controls(lngCont01).Width
                
                For lngCont02 = 0 To frmForm.Controls(lngCont01).Cols - 1
                
                    ctrControlli(lngCont01).lngGrigliaTag = ctrControlli(lngCont01).lngGrigliaTag & Format(frmForm.Controls(lngCont01).ColWidth(lngCont02), "00000")
                
                Next lngCont02
            
            End If
            
        End If
    
    Next lngCont01
    
    If (frmForm.Controls.Count > 0) Then
        For lngCont01 = 0 To (frmForm.Controls.Count - 2)
            For lngCont02 = lngCont01 + 1 To (frmForm.Controls.Count - 1)
                If (ctrControlli(lngCont01).lngLevel > ctrControlli(lngCont02).lngLevel) Then
                    ctrControlloTemp = ctrControlli(lngCont01)
                    ctrControlli(lngCont01) = ctrControlli(lngCont02)
                    ctrControlli(lngCont02) = ctrControlloTemp
                End If
            Next lngCont02
        Next lngCont01
    End If
    
End Sub
' ========================================================================================= '
' Routine: Ridimensiona                                                  Creata: 23/11/2002 '
' Visibilità: Public                                         Autore: Massimiliano Mazzoli   '
' Descrizione: Ridimensiona i controlli posizionati su di una Form                          '
' ========================================================================================= '
Public Sub Ridimensiona(ByRef frmForm As Form)
    
    Dim dblHeightChange As Double
    Dim dblWidthChange As Double
    Dim lngCont01 As Long
    Dim strTag As String
    Dim lngIndex As Long
    Dim dblSurplus As Double
    Dim lngCont02 As Long
                
On Error GoTo ERR_SET
                
    If (frmForm.Controls.Count = 0) Then
        Exit Sub
    End If
    
    If (frmForm.WindowState = vbMinimized) Then
        Exit Sub
    End If
        
    If (frmForm.WindowState = vbNormal) Then
    
        If ((lngFormHeightMin > 0) And (frmForm.Height < lngFormHeightMin)) Then
            frmForm.Height = lngFormHeightMin
        End If
    
        If ((lngFormWidthMin > 0) And (frmForm.Width < lngFormWidthMin)) Then
            frmForm.Width = lngFormWidthMin
        End If
    
    End If

    Call LockWindowUpdate(frmForm.hWnd)
   
    
    For lngCont01 = 0 To UBound(ctrControlli)
    
        lngIndex = ctrControlli(lngCont01).lngIndex
        
        dblHeightChange = frmForm.Controls(lngIndex).Container.Height / ctrControlli(lngCont01).lngContainerHeight
        dblWidthChange = frmForm.Controls(lngIndex).Container.Width / ctrControlli(lngCont01).lngContainerWidth
        
        strTag = TagEstrai(frmForm.Controls(lngIndex).Tag, "rs=")
        
        strTag = UCase(strTag)

        If (strTag = "") Then
        
            strTag = "L,T,W,H"
        
        End If

        If (TypeOf frmForm.Controls(lngIndex) Is Line) Then
        
            If (CBool(InStr(strTag, "L"))) Then
            
                frmForm.Controls(lngIndex).X1 = (ctrControlli(lngCont01).lngLeft * dblWidthChange)
      
            End If
        
            If (CBool(InStr(strTag, "T"))) Then
            
                frmForm.Controls(lngIndex).Y1 = (ctrControlli(lngCont01).lngTop * dblHeightChange)
            
            End If
        
            If (CBool(InStr(strTag, "W"))) Then
            
                frmForm.Controls(lngIndex).X2 = (ctrControlli(lngCont01).lngWidth * dblWidthChange)
                            
            End If
            
            If (CBool(InStr(strTag, "H"))) Then
        
                frmForm.Controls(lngIndex).Y2 = (ctrControlli(lngCont01).lngHeight * dblHeightChange)
                    
            End If
        
        Else
            
           ' On Error Resume Next
    
            If (CBool(InStr(strTag, "L"))) Then
            
                If (frmForm.Controls(lngIndex).Left < 0) Then
                    
                    frmForm.Controls(lngIndex).Left = ((ctrControlli(lngCont01).lngLeft * dblWidthChange) + ssTabOffset)
                
                Else
                    
                    frmForm.Controls(lngIndex).Left = (ctrControlli(lngCont01).lngLeft * dblWidthChange)
                
                End If
            
            End If

            If (CBool(InStr(strTag, "T"))) Then
            
                frmForm.Controls(lngIndex).TOp = (ctrControlli(lngCont01).lngTop * dblHeightChange)
            
            End If
            
            If (CBool(InStr(strTag, "W"))) Then
            
                frmForm.Controls(lngIndex).Width = (ctrControlli(lngCont01).lngWidth * dblWidthChange)
                            
            End If
            
            If (CBool(InStr(strTag, "H"))) Then
        
                frmForm.Controls(lngIndex).Height = (ctrControlli(lngCont01).lngHeight * dblHeightChange)
                    
            End If
        
            If (CBool(InStr(strTag, "G0")) Or CBool(InStr(strTag, "G1"))) Then
                
                If (frmForm.Controls(lngIndex).Width >= ctrControlli(lngCont01).lngGrigliaWidth) Then
                    
                    dblSurplus = (frmForm.Controls(lngIndex).Width - ctrControlli(lngCont01).lngGrigliaWidth) / frmForm.Controls(lngIndex).Cols
                
                Else
                    
                    dblSurplus = 0
                
                End If
                
                For lngCont02 = 1 To frmForm.Controls(lngIndex).Cols - 1
                    
                    frmForm.Controls(lngIndex).ColWidth(lngCont02) = Val(Mid(ctrControlli(lngCont01).lngGrigliaTag, lngCont02 * 5 + 1, 5)) + dblSurplus
                
                Next lngCont02
                        
                If CBool(InStr(strTag, "G1")) Then
                    
                    Call GrigliaAllineaUltimaColonna(frmForm.Controls(lngIndex))
                
                End If
            End If
        
        End If
    
    Next lngCont01

    Call LockWindowUpdate(0)
    
ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_SET:
    Resume Next
End Sub

' ========================================================================================= '
' Routine: TagEstrai                                                     Creata: 18/11/2002 '
' Visibilità: Private                                        Autore: Massimiliano Mazzoli   '
' Descrizione: Funzione che estrae il valore di una chiave contenuta nella proprietà Tag    '
' ========================================================================================= '
Private Function TagEstrai(ByVal strTag As String, _
                           ByVal strChiave As String, _
                           Optional ByVal strChiaveFine As String = " ") As Variant

    TagEstrai = ""
    
    Dim strTemp() As String
    Dim intPos01 As Integer
    Dim intPos02 As Integer
    Dim intCont01 As Integer
    Dim strTemp01 As String
    
    intCont01 = 0
    
    strTag = strTag & strChiaveFine
    
    ReDim strTemp(intCont01)
    
    intPos01 = InStr(1, strTag, strChiave, vbTextCompare)
    
    While ((intPos01 < Len(strTag)) And (intPos01 > 0))
            
        intPos02 = InStr(intPos01, strTag, strChiaveFine, vbTextCompare)
        
        If (intPos02 >= intPos01) Then
            
            ReDim Preserve strTemp(intCont01)
    
            strTemp(intCont01) = Mid(strTag, intPos01 + Len(strChiave), intPos02 - (intPos01 + Len(strChiave)))
        
            intPos01 = InStr(intPos01 + 1, strTag, strChiave, vbTextCompare)
        
            intCont01 = intCont01 + 1
        Else
        
            intPos01 = 0
        
        End If
    
    Wend
    
    If (UBound(strTemp)) = 0 Then
        
        TagEstrai = strTemp(0)
        
    Else
    
        TagEstrai = strTemp()
    
    End If
    
End Function

' ========================================================================================= '
' Routine: TagEstrai                                                     Creata: 18/11/2002 '
' Visibilità: Private                                        Autore: Massimiliano Mazzoli   '
' Descrizione: Funzione che assegna il valore ad una chiave contenuta nella proprietà Tag   '
' ========================================================================================= '
Private Sub TagAssegna(ByRef objOggetto As Object, _
                       ByVal strStringa As String, _
                       ByVal strValore As String, _
                       Optional ByVal strChiaveFine As String = " ")

    Dim intPos01 As Integer
    Dim intPos02 As Integer
    Dim strTemp As String
    
    strTemp = objOggetto.Tag
    intPos01 = InStr(1, objOggetto.Tag, strStringa, vbTextCompare)
        
    If (intPos01 > 0) Then
        If (intPos01 = 1) Then
            objOggetto.Tag = ""
        Else
            objOggetto.Tag = Left(strTemp, intPos01 - 1)
        End If
        objOggetto.Tag = objOggetto.Tag & strStringa & strValore & strChiaveFine
        intPos02 = InStr(intPos01, objOggetto.Tag, strChiaveFine, vbTextCompare)
        objOggetto.Tag = objOggetto.Tag & Right(strTemp, Len(strTemp) - intPos02)
    Else
        objOggetto.Tag = objOggetto.Tag & strChiaveFine & strStringa & strValore & strChiaveFine
    End If
    
End Sub

' ========================================================================================= '
' Routine: LivelloConta                                                  Creata: 18/11/2002 '
' Visibilità: Private                                        Autore: Massimiliano Mazzoli   '
' Descrizione: Determina il livello di un controllo, cioè se è contenuto in altri controlli '
'              o solo alla Form
' ========================================================================================= '
Private Function LivelloConta(objOggetto As Object) As Integer
  
    If TypeOf objOggetto.Container Is Form Then
        
        LivelloConta = 1
    
    Else
        
        LivelloConta = LivelloConta(objOggetto.Container) + 1
    
    End If

End Function

' ========================================================================================= '
' Routine: GrigliaAllineaUltimaColonna                                  Creata: 18/11/2002  '
' Visibilità: Private                                        Autore: Massimiliano Mazzoli   '
' Descrizione: Allinea l'ultima colonna della griglia alla destra della stessa              '
' ========================================================================================= '
Private Sub GrigliaAllineaUltimaColonna(ByRef ctrGriglia As Control)

    Dim intCont01 As Integer
    Dim lngColonneLarghezzaTotale As Long
    
    intCont01 = 0
    lngColonneLarghezzaTotale = 0
    
    For intCont01 = 0 To ctrGriglia.Cols - 2
        
        lngColonneLarghezzaTotale = lngColonneLarghezzaTotale + ctrGriglia.ColWidth(intCont01)
        
    Next intCont01
    
    ctrGriglia.ColWidth(ctrGriglia.Cols - 1) = ctrGriglia.Width - lngColonneLarghezzaTotale - 110

End Sub

' ========================================================================================= '
' Routine: FormCentra                                                    Creata: ??/??/???? '
' Visibilità: Public                                                     Autore: Ignoto     '
' Descrizione: Centra la Form passata al centro dello schermo                               '
'                                                                                           '
' ========================================================================================= '
Public Sub FormCentra(ByRef frmForm As Form)

    ' Centra una form nello schermo
    frmForm.TOp = (Screen.Height - frmForm.Height) / 2
    frmForm.Left = (Screen.Width - frmForm.Width) / 2
    
End Sub
