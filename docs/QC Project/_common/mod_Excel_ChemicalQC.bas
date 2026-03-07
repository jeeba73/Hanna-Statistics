Attribute VB_Name = "mod_Excel_ChemicalQC"
Option Explicit
Public SHEET_NAME As String

Private SettingName As String

Public sPath As String
Private maxRecord As Integer
Private Const bIntro As Boolean = True
Private LastRow As Integer

Private MeasurementUnit As String
Private STDCount As String
Private Virgola As Integer
Private UserDecimal As String
Private STD() As String
Private ph() As String
Private Family() As String
Private FamilyClean() As String
Private pHNumber As Integer
Private MeterNumber As Integer
Private OtherCode() As String
Private OtherParameterFormula() As String
Private OtherCount As Integer


Private Sub SetUnit()
Dim Code As String
Dim Recipe As String
OtherCount = 0


        Family() = FamilyClean
    
        STDCount = GetSettingData(SettingName, "Graph QC", "STDCount", 0)
        Virgola = GetSettingData(SettingName, "Code Information", "Decimal", 0)
        UserDecimal = FormatDecimal(GetSettingData(SettingName, "Code Information", "Decimal", 0))
        MeasurementUnit = GetSettingData(SettingName, "Code Information", "MeasurementUnit", "")
        MeterNumber = GetSettingData(SettingName, "Information QC", "MeterNumber", 0)
        Code = GetSettingData(SettingName, "Code Information", "Code", "")
        Recipe = GetSettingData(SettingName, "Code Information", "Recipe", "")
        
        
        Call GetSTDTable
        Call GetpH
        Call GetOtherCode(Recipe, Code, OtherCount)
        
End Sub

Public Function CopyChemicalQCData(Optional ByVal FileName As String) As Boolean
    Dim i As Long
    Dim a As Integer
    Dim rc As Boolean
    Dim strPath As String
    
    On Error GoTo ERR_SIT
    rc = True
    SettingName = FileName
    
    strPath = USER_PATH
    
    
    If FileExists(USER_TEMP_PATH & SettingName) Then
        USER_PATH = USER_TEMP_PATH
    ElseIf FileExists(USER_DATA_PATH & SettingName) Then
         USER_PATH = USER_DATA_PATH
    Else
        PopupMessage 2, "file doesn't exist", , True, SettingName
        rc = False
        GoTo ERR_END
    End If
    
    
    '---------------------------
    ' set excel page
    '---------------------------
    Call SetUnit
    Call FormatPage
    
     
    '---------------------------
    ' copy data in excel page
    '---------------------------
    
    
    Call SetInformationQC(i)
    
    
    Call SetReadingData(i)
    
    
    Call SetEvaluationData(i)

    CloseSettingDataFile
ERR_END:
    On Error GoTo 0
    CopyChemicalQCData = rc
    USER_PATH = strPath
    Exit Function
ERR_SIT:
    rc = False
    MsgBox a & Err.Description
    Resume Next
End Function

Private Sub SetInformationQC(ByRef Riga As Long)
Dim i As Integer
Dim sString As String
Dim rc As Boolean

On Error GoTo ERR_SET:


    Call AddValue(2, 2, "Lot", True, True)
    Call AddValue(3, 2, GetSettingData(SettingName, "Information QC", "Text10", ""), , True)
    Call AddValue(2, 3, "Hanna SFG Code", True, True)
    Call AddValue(3, 3, GetSettingData(SettingName, "Information QC", "Text11", ""), , True)
    Call AddValue(2, 4, "Recipe", True, True)
    Call AddValue(3, 4, GetSettingData(SettingName, "Information QC", "Text15", ""), , True)
    Call AddValue(2, 5, "Preparation Week", True, True)
    Call AddValue(3, 5, GetSettingData(SettingName, "Information QC", "Text121", ""), , True)

    Call AddValue(6, 2, "Description", True)
    Call AddValue(7, 2, GetSettingData(SettingName, "Information QC", "Text12", ""))
    Call AddValue(6, 3, "Exp", True)
    Call AddValue(7, 3, GetSettingData(SettingName, "Information QC", "Text13", ""))
    Call AddValue(6, 4, "Line", True)
    Call AddValue(7, 4, GetSettingData(SettingName, "Information QC", "Text14", ""))
    
    If OtherCount > 0 Then
    
        Call AddValue(2, 7, "Hanna SFG Code", True, True)
        Call AddValue(3, 7, "Parameter Formula", True, True)
        
        For i = 0 To OtherCount - 1
            Call AddValue(2, 8 + i, OtherCode(i))
            Call AddValue(3, 8 + i, OtherParameterFormula(i))
        Next
    
    End If
    
    
    Call AddValue(6, 7, "QC Department", True)
    Call AddValue(7, 7, GetSettingData(SettingName, "Information QC", "Text131", ""))
    Call AddValue(6, 8, "Registration Book", True)
    Call AddValue(7, 8, GetSettingData(SettingName, "Information QC", "Text132", ""))
    
    Call AddValue(6, 9, "QC Type", True)
    Call AddValue(7, 9, GetSettingData(SettingName, "Information QC", "Text130", ""))
    
    
    
    
    
    Call AddValue(8, 2, "Ref Weight.", True)
    Call AddValue(9, 2, GetSettingData(SettingName, "Information QC", "Text16", ""))
    Call AddValue(8, 3, "Min (mg)", True)
    Call AddValue(9, 3, GetSettingData(SettingName, "Information QC", "Text17", ""))
    Call AddValue(8, 4, "Max (mg)", True)
    Call AddValue(9, 4, GetSettingData(SettingName, "Information QC", "Text18", ""))

    Call AddValue(13, 2, "REAGENT SET 1", True)
    Call AddValue(13, 8, "REAGENT SET 2", True)
    For i = 1 To 5
        Call AddValue(14, 1 + i, "Reagent " & Chr$(64 + i) & " Lot", True)
        Call AddValue(15, 1 + i, GetSettingData(SettingName, "Information QC", "Text1" & i + 10, ""))
        Call AddValue(16, 1 + i, "Expiration " & Chr$(64 + i), True)
        Call AddValue(17, 1 + i, GetSettingData(SettingName, "Information QC", "Text1" & i + 15, ""))
        
        Call AddValue(14, 7 + i, "Reagent " & Chr$(64 + i) & " Lot", True)
        Call AddValue(15, 7 + i, GetSettingData(SettingName, "Information QC", "Text1" & i + 44, ""))
        Call AddValue(16, 7 + i, "Expiration " & Chr$(64 + i), True)
        Call AddValue(17, 7 + i, GetSettingData(SettingName, "Information QC", "Text1" & i + 49, ""))
        
    Next
    
Riga = 17

    Call AddValue(Riga + 2, 2, "Reagent Range Min", True)
    Call AddValue(Riga + 3, 2, GetSettingData(SettingName, "Information QC", "Text19", "") & " " & MeasurementUnit)
    Call AddValue(Riga + 2, 3, "Reagent Range Max", True)
    Call AddValue(Riga + 3, 3, GetSettingData(SettingName, "Information QC", "Text110", "") & " " & MeasurementUnit)

Riga = Riga + 4


    
    Call AddValue(Riga + 2, 2, "Prep. Operator", True)
    Call AddValue(Riga + 3, 2, GetSettingData(SettingName, "Information QC", "Text122", ""))
    Call AddValue(Riga + 2, 3, "First day Prod.", True)
    Call AddValue(Riga + 3, 3, GetSettingData(SettingName, "Information QC", "Text123", ""))
    Call AddValue(Riga + 2, 4, "Last day Prod.", True)
    Call AddValue(Riga + 3, 4, GetSettingData(SettingName, "Information QC", "Text124", ""))
    Call AddValue(Riga + 2, 5, "Machine", True)
    Call AddValue(Riga + 3, 5, GetSettingData(SettingName, "Information QC", "Text125", ""))
    
    Call AddValue(Riga + 2, 8, "Old Lot A", True)
    Call AddValue(Riga + 3, 8, GetSettingData(SettingName, "Information QC", "Text126", ""))
    Call AddValue(Riga + 4, 8, "Expiration A", True)
    Call AddValue(Riga + 5, 8, GetSettingData(SettingName, "Information QC", "Text127", ""))
    
    Call AddValue(Riga + 2, 9, "Old Lot B", True)
    Call AddValue(Riga + 3, 9, GetSettingData(SettingName, "Information QC", "Text128", ""))
    Call AddValue(Riga + 4, 9, "Expiration B", True)
    Call AddValue(Riga + 5, 9, GetSettingData(SettingName, "Information QC", "Text129", ""))
    
    Call AddValue(Riga + 2, 10, "Old Lot C", True)
    Call AddValue(Riga + 3, 10, GetSettingData(SettingName, "Information QC", "Text158", ""))
    Call AddValue(Riga + 4, 10, "Expiration C", True)
    Call AddValue(Riga + 5, 10, GetSettingData(SettingName, "Information QC", "Text159", ""))

  
    Call AddValue(Riga + 2, 10, "Old Lot D", True)
    Call AddValue(Riga + 3, 10, GetSettingData(SettingName, "Information QC", "Text160", ""))
    Call AddValue(Riga + 4, 10, "Expiration D", True)
    Call AddValue(Riga + 5, 10, GetSettingData(SettingName, "Information QC", "Text161", ""))


    
    
    
    Riga = Riga + 8
    
    Call AddValue(Riga, 2, "Meter", True, True)
    
    Riga = Riga + 1
    
    ReDim Family(MeterNumber)
    
    For i = 1 To MeterNumber
    
        Family(i - 1) = GetSettingData(SettingName, "Information QC", "Text1" & 31 + i * 2, "")
        Call AddValue(Riga, i * 2, "Meter " & i & " Family", True)
        Call AddValue(Riga, 1 + i * 2, "Meter " & i & " Code", True)
        Call AddValue(Riga + 1, i * 2, Family(i - 1))
        Call AddValue(Riga + 1, 1 + i * 2, GetSettingData(SettingName, "Information QC", "Text1" & 32 + i * 2, ""))
    Next
        
    Call AddValue(Riga + 2, 2, "ph Meter ", True)
    Call AddValue(Riga + 2, 3, "Description", True)
    Call AddValue(Riga + 3, 2, GetSettingData(SettingName, "Information QC", "Text141", ""))
    Call AddValue(Riga + 3, 3, GetSettingData(SettingName, "Information QC", "Text142", ""))
    
    Call AddValue(Riga + 2, 4, "Turbid. Meter ", True)
    Call AddValue(Riga + 2, 5, "Description", True)
    Call AddValue(Riga + 3, 4, GetSettingData(SettingName, "Information QC", "Text143", ""))
    Call AddValue(Riga + 3, 5, GetSettingData(SettingName, "Information QC", "Text144", ""))
    
    Call AddValue(Riga + 2, 6, "Spectr. Meter ", True)
    Call AddValue(Riga + 2, 7, "Description", True)
    Call AddValue(Riga + 3, 6, GetSettingData(SettingName, "Information QC", "Text156", ""))
    Call AddValue(Riga + 3, 7, GetSettingData(SettingName, "Information QC", "Text157", ""))
       
       
ERR_END:
    On Error GoTo 0
    

    Riga = Riga + 2
    
    Exit Sub
ERR_SET:
    MsgBox Err.Description
    Resume Next
 
End Sub


Private Sub SetReadingData(ByRef Riga As Long)

Dim i As Integer
Dim t As Integer

Dim PrimaRiga As Long
Dim Ultimariga As Long
Dim MyRows As Long
Dim MyCols As Long
Dim COLOR_FORE As OLE_COLOR

On Error GoTo ERR_SET:

Riga = Riga + 4


    Call AddValue(Riga, 2, "Standard", True, True)
    Call AddValue(Riga, 3, MeasurementUnit)
    Call AddValue(Riga, 9, "Tolerance", True, True)
    Call AddValue(Riga, 14, "pH", True, True)
    
    
'Riga = Riga + 1

    Call AddValue(Riga + 1, 2, "STD Number", True)
    Call AddValue(Riga + 1, 3, "STD Value", True)
    Call AddValue(Riga + 1, 4, "STD Min", True)
    Call AddValue(Riga + 1, 5, "STD Max", True)
    
    Call AddValue(Riga + 1, 7, "STD MR", True)
    Call AddValue(Riga + 2, 7, GetSettingData(SettingName, "Code Information", "STDMR", ""))
    
    
    
    
    Call AddValue(Riga + 1, 9, "Fixed", True)
    Call AddValue(Riga + 2, 9, GetSettingData(SettingName, "Code Information", "Fixed", ""))
    Call AddValue(Riga + 1, 10, "AndOr", True)
    Call AddValue(Riga + 2, 10, GetSettingData(SettingName, "Code Information", "AndOr", ""))
    Call AddValue(Riga + 1, 11, "Percentage", True)
    Call AddValue(Riga + 2, 11, GetSettingData(SettingName, "Code Information", "Percentage", "") & "%")
    Call AddValue(Riga + 1, 12, "QCRestriction", True)
    Call AddValue(Riga + 2, 12, GetSettingData(SettingName, "Code Information", "QCRestriction", "") & "%")
    
    
    Call AddValue(Riga + 1, 14, "pH Value", True)
    Call AddValue(Riga + 1, 15, "pH Min", True)
    Call AddValue(Riga + 1, 16, "pH Max", True)
    
     For i = 1 To 3
        Call AddValue(Riga + 1 + i, 14, ph(i, 0))
        Call AddValue(Riga + 1 + i, 15, ph(i, 1))
        Call AddValue(Riga + 1 + i, 16, ph(i, 2))
    Next
           
    For i = 1 To STDCount
        Call AddValue(Riga + 1 + i, 2, STD(i, 0))
        Call AddValue(Riga + 1 + i, 3, STD(i, 1))
        Call AddValue(Riga + 1 + i, 4, STD(i, 2))
        Call AddValue(Riga + 1 + i, 5, STD(i, 3))
    Next
    
    

        
    Riga = Riga + 3 + i
    
    
    Dim Rows As Long
    

    MyRows = GetSettingData(SettingName, "Reading QC", "Grd2 Rows", 1)
    Rows = CLng(MyRows)
    
    If Rows = 1 Then
        Call CheckRows(Rows, SettingName, USER_PATH)
        MyRows = Rows
    End If
     
    MyCols = GetSettingData(SettingName, "Reading QC", "Grd2 Cols", 1)


    Dim bNonSelezionato     As Boolean
    Dim Standard            As Integer
    Dim Test                As Integer
    Dim Meter               As Integer
    Dim Value               As String
    Dim QCType              As String
    Dim b                   As Integer
    If MyRows > 1 Then
    
        Call AddValue(Riga + 1, 2, "Readings Table", True, True)
    
        Riga = Riga + 2
    
        If MyCols > 1 Then
            For i = 0 To MyRows - 1
                t = 0
                
                QCType = GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & 4, 0)
                If QCType = "" Then GoTo contReading:
                            
                For b = 1 To MyCols - 1
                    If b >= MyCols - 12 And b < MyCols - 2 Then
                        GoTo continua:
                    End If
                    t = t + 1
                    COLOR_FORE = GetSettingData(SettingName, "Reading QC", "Grd2 Fore Row" & i & " Col" & b, vbBlack)
                    If i = 0 Then
                        Call AddValue(Riga + i, 1 + t, GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & b, ""), True)
                    ElseIf t = 11 Or t = 12 Or t = 13 Or t = 14 Then
                        Meter = t - 10
                        If Meter > MeterNumber Then
                             Call AddValue(Riga + i, 1 + t, "/")
                        Else

                            Standard = CInt(GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & 1, 0))
                            Test = GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " STD Test", 0) 'GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & 3, 0)
                            bNonSelezionato = CheckSelezionato(Standard, Test, Meter)
                            Value = GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & b, "")
                            If (InStr(QCType, "P")) = 0 Then
                                Call AddValue(Riga + i, 1 + t, Value, , , COLOR_FORE, True)
                            Else
                                If Value <> "" Then
                                    Call AddValue(Riga + i, 1 + t, Value, , , COLOR_FORE, bNonSelezionato)
                                Else
                                    GoTo Add
                                End If
                            End If
                        End If
                    Else
Add:
                        Call AddValue(Riga + i, 1 + t, GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & b, ""), , , COLOR_FORE)
           
                    End If
continua:
                Next
contReading:
            Next
        End If
    End If
    
   Riga = Riga + i + 3
   

ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_SET:
    MsgBox Err.Description
    Resume Next
End Sub
Private Function CheckSelezionato(ByVal Standard As Integer, ByVal Test As Integer, ByVal Meter As Integer) As Boolean
Dim rc As Boolean
    rc = GetSettingData(SettingName, "Graph QC", "Standard " & Standard & " Test " & Test & " Meter " & Meter & " Selected", "TRUE")
    CheckSelezionato = Not (rc)
End Function

Private Sub SetEvaluationData(ByRef Riga As Long)
Dim i As Integer
Dim t As Integer
Dim PrimaRiga As Long
Dim Ultimariga As Long
Dim nRows As Long
Dim nCols As Long
Dim COLOR_FORE As OLE_COLOR
Dim MaxTest



    '  tabella Results
    
    If STDCount = 0 Then Exit Sub
    
   
       
    MyChemicalQC = MyChemicalQCClean
    

    
    'Call AddValue(Riga, 7, "Average", True, True)
    
    Riga = Riga - 1
    
    For i = 1 To STDCount
    
 
        t = CInt(STD(i, 0))

        Call AddValue(Riga + i * 4, 2, "Standard # " & t, True)
        
        MyChemicalQC.STDtest(t).MaxReadings = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Total Readings", 0)
        MyChemicalQC.STDtest(t).SelReadings = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Selected Readings", 0)
        MyChemicalQC.STDtest(t).NumTest = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Total Tests", 0)
        MyChemicalQC.STDtest(t).SelTest = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Selected Tests", 0)
        MyChemicalQC.STDtest(t).TotalMean = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Total Average", 0)
        MyChemicalQC.STDtest(t).SelecMean = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Selected Average", 0)
         
        Call AddValue(Riga + i * 4 + 1, 2, "Total Tests", True)
        Call AddValue(Riga + i * 4 + 2, 2, CStr(MyChemicalQC.STDtest(t).NumTest))

        Call AddValue(Riga + i * 4 + 1, 3, "Total Readings", True)
        Call AddValue(Riga + i * 4 + 2, 3, CStr(MyChemicalQC.STDtest(t).MaxReadings))
        
        Call AddValue(Riga + i * 4 + 1, 4, "Total Average", True)
        Call AddValue(Riga + i * 4 + 2, 4, MyChemicalQC.STDtest(t).TotalMean & " " & MeasurementUnit)
        
        Call AddValue(Riga + i * 4 + 1, 5, "Selected Tests", True)
        Call AddValue(Riga + i * 4 + 2, 5, CStr(MyChemicalQC.STDtest(t).SelTest))
        
        Call AddValue(Riga + i * 4 + 1, 6, "Selected Readings", True)
        Call AddValue(Riga + i * 4 + 2, 6, CStr(MyChemicalQC.STDtest(t).SelReadings))
        
        Call AddValue(Riga + i * 4 + 1, 7, "Selected Average", True)
        Call AddValue(Riga + i * 4 + 2, 7, MyChemicalQC.STDtest(t).SelecMean & " " & MeasurementUnit)
        
    Next
    
    MyChemicalQC = MyChemicalQCClean

    Riga = Riga + i * 4 + 4
 
    nRows = GetSettingData(SettingName, "Evaluation QC", "Results Grid Rows", 0)
    nCols = GetSettingData(SettingName, "Evaluation QC", "Results Grid Cols", 0)
    
    If nRows = 0 Then Exit Sub
    
    
    Call AddValue(Riga + 1, 2, "Specifications", True, True)
    
    Riga = Riga + 2
    
    
    For i = 0 To nRows - 1
        For t = 1 To nCols - 1
            COLOR_FORE = GetSettingData(SettingName, "Evaluation QC", "Results Grid Standard (" & i & ") Forecolor " & t, vbBlack)
            If t = 4 Or t = 5 Or t = 6 Then
            ElseIf t = 7 Then
                If i = 0 Then
                    Call AddValue(Riga + i, 1 + t - 3, GetSettingData(SettingName, "Evaluation QC", "Results Grid Standard (" & i & ")  Column " & t, ""), True, True)
                Else
                
                    Call AddValue(Riga + i, 1 + t - 3, GetSettingData(SettingName, "Evaluation QC", "Results Grid Standard (" & i & ")  Column " & t, ""), , , COLOR_FORE)
                End If
            
            Else
                If i = 0 Then
                    Call AddValue(Riga + i, 1 + t, GetSettingData(SettingName, "Evaluation QC", "Results Grid Standard (" & i & ")  Column " & t, ""), True)
                Else
                
                    Call AddValue(Riga + i, 1 + t, GetSettingData(SettingName, "Evaluation QC", "Results Grid Standard (" & i & ")  Column " & t, ""), , , COLOR_FORE)
                End If
            End If
            
           
        Next
    Next

    Riga = Riga + 3 + i
    
    Call AddValue(Riga, 2, "Validation Date", True)
    Call AddValue(Riga + 1, 2, GetSettingData(SettingName, "Close QC", "Validation Date", ""))
    Call AddValue(Riga, 3, "by", True)
    Call AddValue(Riga + 1, 3, GetSettingData(SettingName, "Close QC", "Operator", ""))
    
    
    

    


End Sub


Private Sub GetSTDTable()

Dim t As Integer
Dim k As Integer
Dim rc As Boolean

    If STDCount = 0 Then
        Exit Sub
    End If
        
        ReDim STD(STDCount, 3) As String
        For t = 1 To STDCount
            STD(t, 0) = GetSettingData(SettingName, "Graph QC", "STDNumber" & t, "")
            STD(t, 1) = GetSettingData(SettingName, "Graph QC", "STDValue" & t, "")
            STD(t, 2) = GetSettingData(SettingName, "Graph QC", "STDMin" & t, "")
            STD(t, 3) = GetSettingData(SettingName, "Graph QC", "STDMax" & t, "")
            
            If InStr(STD(t, 1), "/") Or STD(t, 1) = "" Then
                STD(t, 1) = 0
            End If
            If InStr(STD(t, 2), "/") Or STD(t, 2) = "" Then
                STD(t, 2) = 0
            End If
            If InStr(STD(t, 3), "/") Or STD(t, 3) = "" Then
                STD(t, 3) = 0
            End If
            CloseSettingDataFile
        Next
    
ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_GET:
    rc = False
    MsgBox Err.Description
    Resume Next
End Sub

Private Function GetpH()
Dim i As Integer
Dim Num As Integer


   ' pHNumber = GetSettingData(SettingName, "Information QC", "pHNumber", 0)
    
    ReDim ph(3, 2) As String
        
    For Num = 1 To 3
        ph(Num, 0) = GetSettingData(SettingName, "Information QC", "pHValue" & Num, "0")
        ph(Num, 1) = GetSettingData(SettingName, "Information QC", "pHMin" & Num, "0")
        ph(Num, 2) = GetSettingData(SettingName, "Information QC", "pHMax" & Num, "0")
    Next
End Function

Public Function GetOtherCode(ByVal Recipe As String, ByVal Code As String, ByRef OtherCount As Integer)
Dim i As Integer

On Error GoTo ERR_GET:

    With dbTabCode
        .filter = ""
        .filter = "Recipe='" & Trim(Recipe) & "'"
        If .EOF Then
        Else
            .MoveFirst
            ReDim OtherCode(.RecordCount) As String
            ReDim OtherParameterFormula(.RecordCount) As String
            
            For i = 1 To .RecordCount
                'If Trim(!Code) = Trim(Code) Then
                'Else
                    If OtherCount > 0 Then
                        ' controllo se non lo ho già inserito
                        If OtherCode(OtherCount - 1) <> Trim(!Code) Then
Add:
                            OtherCode(OtherCount) = Trim(!Code)
                            OtherParameterFormula(OtherCount) = IIf(IsNull(Trim(!ParameterFormula)), "", Trim(!ParameterFormula))
                            OtherCount = OtherCount + 1
                        End If
                    Else
                        GoTo Add:
                    End If
                'End If
                .MoveNext
            Next
            ReDim Preserve OtherCode(OtherCount) As String
    
        End If

    End With
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_GET:
    MsgBox Err.Description
    Resume Next
    
End Function
