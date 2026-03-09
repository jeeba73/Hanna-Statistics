Attribute VB_Name = "mod_word"
Option Explicit
Private UsedVariables() As String
Public Stringa_metodo
Public mevar
Public StrWord
Public SaveString As String
Public bSeStampa As Boolean

Public USE_FIRMA As Boolean

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'                           inserisce le variabili
'                               nel foglio Word
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function LetPrint(ByVal NumReport As String, Optional ByVal bPrint As Boolean = True) As Boolean
    If LoadCertificato Then
        LetPrint = F_PRINT.DoShow(NumReport, bPrint)
    End If
End Function

Public Function CreateWord(ByRef WordApp As Object) As Boolean
    Dim rc As Boolean
    On Error GoTo ERR_CREATE
    rc = True

    Set WordApp = GetObject(, "Word.Application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("Word.Application")
        If WordApp Is Nothing Then
        PopupMessage 2, "Impossibile aprire Microsoft Word", , True
        rc = False
        End If
    End If
    On Error GoTo 0
    CreateWord = rc
    Exit Function
ERR_CREATE:

    If Err.NUMBER = 429 Then
        Err.Clear
        rc = True
       ' MsgBox err.NUMBER
        Resume Next
    End If
    rc = False
    MessageInfoTime = 2000
    PopupMessage 2, Err.Description, , True, "Microsoft Word"
    
    Resume Next
End Function

Public Function CreaReport(ByVal WordApp As Object, ByVal NumReport As String, Optional ByVal StringReport As String = "Report n.", _
                           Optional ByVal LoadPath As String, Optional ByVal SavePath As String) As Boolean
    
    '-------------------------------------------
    '   crea il report in formato Word
    '-------------------------------------------
    
    Dim WordDoc As Object
    Dim docOpen As String
    
    Dim sPath As String
    Dim lPath As String
    Dim rc As Boolean
    On Error GoTo ERR_MAKE
    
    

    rc = True
    
    docOpen = DOC_NAME
    
    ReDim UsedVariables(0)
    
    sPath = CheckSavePath(SavePath)
    lPath = CheckSavePath(LoadPath)
    Debug.Print lPath
    
    
     If FileExists(lPath & docOpen) Then
    Else
        FileCopy App.Path & "\bin\" & docOpen, lPath & docOpen
        DoEvents
    End If
checkFile:
    
    If FileIsOpen(lPath & docOpen) Then

        
        MessageInfoTime = 2000
        PopupMessage 2, "Attenzione il file Template " & docOpen & " risulta aperto." & vbCrLf & "impossibile procedere con la creazione del Reprot/Certificato", , True
        MessageInfoTime = 2000
        
         PopupMessage 2, "Chiudere tutti i file Word quindi ristampare il documento..."
        GoTo ERR_END
        rc = False
    
    Else
    
    End If
    

    Set WordDoc = WordApp.Documents.Open(lPath & docOpen)
    
    Call SetInVariables(WordDoc)
    DoEvents
    
    
   ' If USE_FIRMA Then
      ' Call AddFirma(WordDoc, MyTaratura.Operatore, ResponsabileLaboratorio)
   ' End If
    
   
    
    SaveString = sPath & StringReport & ".doc"

   
    Dim bWord As Boolean
    
    bWord = GetSetting(App.Title, "REPORT", "CREA WORD", False)

    If bStampanteOK Then
    
        If (bWord) Then
        
            WordDoc.SaveAs SaveString & ".doc"
            
        Else

            Call SetPDFPrinter("", SaveString)

            WordDoc.PrintOut
        End If
    Else
        WordDoc.SaveAs SaveString & ".doc"
    End If

    
    
    '-------------------------------------------
    '   stampa se richiesto
    '-------------------------------------------
    If bSeStampa Then WordDoc.PrintOut
    DoEvents
    
   
 
    
ERR_END:
    
    CreaReport = rc
    Set WordDoc = Nothing
    On Error GoTo 0
    Exit Function
ERR_MAKE:
    MsgBox Err.Description
    rc = False
    Resume Next
    
End Function
    
  
Public Function SetInVariables(WordDoc As Object)
   Dim i As Integer
   Dim NewResult As String
   For i = 1 To WordDoc.fields.Count
   NewResult = GetNewResult(WordDoc.fields(i), WordDoc)
        If NewResult <> "" Then
            WordDoc.fields(i).Result.Text = NewResult
        End If
    Next
End Function

Private Function GetNewResult(wField, WordDoc) As String
    Dim StopPos As Long
    Dim Variable As String
    Dim UsedVariable As String
    Dim VariableValue As String
    Dim VarR As String
    
    
    On Error GoTo ERR_RESULT
    
    StopPos = InStrRev(wField.Code, "\*")
    
    If StopPos = 0 Then Exit Function
    
    Variable = Left(wField.Code, StopPos - 3)
    Variable = Right(Variable, Len(Variable) - 14)
        
        VarR = LCase(Left(Variable, 4))
        
    If VarR <> "" Then

        Call SetVariable(VarR, VariableValue, Variable)
        AddNewVariable Variable, VariableValue
        GetNewResult = VariableValue
    End If
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_RESULT:
    MsgBox Err.Description
    'MsgBox varr
    Resume Next
End Function
Public Sub AddNewVariable(Variable As String, TheValue As String)
Dim ArraySize As Integer
    ArraySize = UBound(UsedVariables)
    ReDim Preserve UsedVariables(ArraySize + 1)
    UsedVariables(ArraySize) = Variable & TheValue
End Sub
Public Function CheckUsedVariable(Variable As String) As Boolean
Dim i As Integer
    For i = 0 To UBound(UsedVariables)
        If Left(UsedVariables(i), Len(Variable)) = Variable Then
            CheckUsedVariable = True
            Exit For
        End If
    Next
End Function
Public Function GetVariableValue(Variable As String) As String
Dim i As Integer
    For i = 0 To UBound(UsedVariables)
        If Left(UsedVariables(i), Len(Variable)) = Variable Then
            GetVariableValue = Right(UsedVariables(i), Len(UsedVariables(i)) - Len(Variable))
            Exit For
        End If
    Next
End Function


Private Function AddPicture(WordDoc As Object) As Boolean


 
    Dim oTable1 As Word.Table
    Dim oPara1 As Word.Paragraph
    
    
    Set oPara1 = WordDoc.Content.Paragraphs.Add(WordDoc.Bookmarks.Item("\endofdoc").Range)
    ' oPara1.range.Text = "Insert New Paragraph"
    oPara1.Format.SpaceAfter = 6
    oPara1.Range.InsertParagraphAfter
    
    Dim PictureLocation As String
    
    PictureLocation = App.Path & "\Prova.jpg"
    WordDoc.Bookmarks.Item("\endofdoc").Range.InlineShapes.AddPicture (PictureLocation)



End Function


Private Function AddFirma(WordDoc As Object, ByVal stringTecnicoName As String, ByVal stringRespName As String) As Boolean


    Dim FirmaTecnicoLocation As String
    Dim FirmaRespLocation As String
    Dim oTable1 As Word.Table
    Dim oPara1 As Word.Paragraph
    Dim sBookmark As String
    Dim mBookmark As Object
    Dim i As Integer

    Set oPara1 = WordDoc.Content.Paragraphs.Add(WordDoc.Bookmarks.Item("\endofdoc").Range)
    ' oPara1.range.Text = "Insert New Paragraph"
    oPara1.Format.SpaceAfter = 6
    oPara1.Range.InsertParagraphAfter
    
    
    FirmaTecnicoLocation = USER_DOCUMENTI & "bin\firme\" & stringTecnicoName & ".jpg" 'App.Path & "\Prova.jpg"
    
    FirmaRespLocation = USER_DOCUMENTI & "bin\firme\" & stringRespName & ".jpg" 'App.Path & "\Prova.jpg"
    
    Debug.Print FirmaTecnicoLocation
    Debug.Print FirmaRespLocation
   

    For Each mBookmark In WordDoc.Bookmarks

        If mBookmark.Name = "Tecnico" Then
            If FileExists(FirmaTecnicoLocation) Then
                WordDoc.Bookmarks.Item("Tecnico").Range.InlineShapes.AddPicture (FirmaTecnicoLocation)
            End If
        End If
        
        If mBookmark.Name = "Responsabile" Then
            If FileExists(FirmaRespLocation) Then
                WordDoc.Bookmarks.Item("Responsabile").Range.InlineShapes.AddPicture (FirmaRespLocation)
            End If
        End If

    Next
   
  

End Function

