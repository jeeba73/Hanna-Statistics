Attribute VB_Name = "mod_word_MatReq"
Option Explicit

Public DOC_NAME As String
Public PrintConfronto As Boolean
Public Societa As String
Public strSocieta As String
Public Indirizzo As String

Public Const CertVirgola = 5
Private NumMyReport As String

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Function OkStampa(ByVal NumMyReport As String, ByVal bSeStampa As Boolean, ByVal FileName As String) As Boolean
    Dim rc As Boolean
    On Error GoTo ERR_SAVE
    rc = True
    '-----------------------------------
    '      Stampooooo
    '-----------------------------------
    
    SettingName = FileName
   
    rc = PrintMe(NumMyReport, bSeStampa)
ERR_END:
    On Error GoTo 0
    OkStampa = rc
    Exit Function
ERR_SAVE:
    rc = False
    Resume ERR_END
End Function

Public Function PrintMe(ByVal NumReport As String, Optional ByVal bPrint As Boolean = True) As Boolean
    '-----------------------------------
    '   imposta routine di stampa
    '-----------------------------------
    Dim rc As Boolean
    On Error GoTo ERR_PRINT
    rc = True
    Call LetPrint(NumReport, bPrint)
ERR_END:
    On Error GoTo 0
    PrintMe = rc
    Exit Function
ERR_PRINT:
    MsgBox Err.Description
    rc = False
    Resume ERR_END
End Function

Public Function LoadCertificato() As Boolean

    Dim sNumReport As String
    Dim rc As Boolean
    Dim Cliente As String
    rc = True
    
    DOC_NAME = ""
    

    DOC_NAME = "MaterialRequisition.docx"
     
    
    
    On Error GoTo ERR_LOAD
  
ERR_END:
    LoadCertificato = rc
    Exit Function
ERR_LOAD:
    rc = False
    Resume ERR_END
End Function

Public Sub SetVariable(ByVal VarR As String, ByRef VariableValue As String, ByVal Variable As String)
    '---------------------------------------------------
    ' VARIABILI DEL REPORT : DIRETTA
    '---------------------------------------------------
    
    Dim numvar As String
    Dim i As Integer
    
    numvar = Right(Variable, Len(Variable) - 4)
    If VarR <> "risu" Then
        Call PrimaPagina(VarR, numvar, VariableValue)
    Else
        Call Risultati(VarR, numvar, VariableValue)
    End If
End Sub

Private Sub Risultati(ByVal VarR As String, ByVal numvar As String, ByRef VariableValue As String)
    '---------------------------------------------------
    '       risultati delle tarature
    '---------------------------------------------------
    Dim i As Integer
    Dim t As Integer
    Dim Rows As Integer
    Dim Value As String
    
    On Error GoTo ERR_RIS
    

    If SettingName = "" Then Exit Sub
    
   
        Rows = GetSettingData(SettingName, "Material Requisition", "Rows", 0)
        
        For i = 1 To Rows
         
                Select Case numvar
                    Case "1"
                        VariableValue = VariableValue & i & vbCrLf
                    Case "2"
                        VariableValue = VariableValue & GetSettingData(SettingName, "Material Requisition", "Grd(" & i & ",1)", "") & vbCrLf
                    Case "3"
                        VariableValue = VariableValue & GetSettingData(SettingName, "Material Requisition", "Grd(" & i & ",2)", "") & vbCrLf
                    Case "4"
                        VariableValue = VariableValue & GetSettingData(SettingName, "Material Requisition", "Grd(" & i & ",3)", "") & vbCrLf
                    Case "5"
                        Value = GetSettingData(SettingName, "Material Requisition", "Grd(" & i & ",4)", "")
                        VariableValue = VariableValue & PadString(Value) & vbCrLf
                    Case "6"
                        VariableValue = VariableValue & GetSettingData(SettingName, "Material Requisition", "Grd(" & i & ",5)", "") & vbCrLf
                    Case "7"
                        VariableValue = VariableValue & GetSettingData(SettingName, "Material Requisition", "Grd(" & i & ",6)", "") & vbCrLf
                End Select

cont:
          
            Next
        
        
ERR_END:
    On Error GoTo 0
    CloseSettingDataFile
    Exit Sub
ERR_RIS:
    MsgBox Err.Description
    Resume ERR_END:
End Sub

'Private Function ValString(ByVal DoubleValue As Double) As String
'    ValString = FormatNumber(DoubleValue, MyVirgola + 2)
'End Function

Private Sub PrimaPagina(ByVal VarR As String, ByVal numvar As String, ByRef VariableValue As String)

    CloseSettingDataFile
    
    
  
    Select Case VarR
        '----------------------
        '        pagina 1
        '----------------------
        Case "mana"
             VariableValue = GetSettingData(SettingName, "Material Requisition", "txDocument(5)", "")
        Case "text"
            '------------------------
            ' specifiche massa
            '------------------------
                
                Select Case numvar
                    Case "0"
                         VariableValue = GetSettingData(SettingName, "Material Requisition", "txDocument(0)", "")
                    Case "1"
                         VariableValue = GetSettingData(SettingName, "Material Requisition", "txDocument(1)", "")
                    Case "2"
                         VariableValue = GetSettingData(SettingName, "Material Requisition", "txDocument(2)", "")
                    Case "3"
                         VariableValue = GetSettingData(SettingName, "Material Requisition", "txDocument(3)", "")
                    Case "4"
                         VariableValue = GetSettingData(SettingName, "Material Requisition", "txDocument(4)", "")

                End Select
           
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    End Select
    
    CloseSettingDataFile
             
                    

End Sub
