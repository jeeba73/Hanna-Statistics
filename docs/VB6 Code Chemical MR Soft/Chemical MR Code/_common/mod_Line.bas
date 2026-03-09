Attribute VB_Name = "mod_Line"
Option Explicit


Public Type DataType
    mel As String
End Type

Public Const strSeparator = ";"


' split strng


Public LineList() As String


Public Function SetLine(ByVal Cmb As ComboBox, Optional bAll As Boolean) As Boolean
Dim i As Integer
    If GetAllLine Then
        With Cmb
            .Clear
            If bAll Then .AddItem "All Lines"
             For i = LBound(LineList) To UBound(LineList)
                .AddItem LineList(i)
             Next
            
             .ListIndex = 0
             If .ListCount = 0 Then Cmb.Visible = False
             If UserLineIndex <= .ListCount And bAll Then .ListIndex = UserLineIndex
        End With
    End If
End Function


Public Function GetAllLine() As Boolean
Dim i As Integer
Dim t As Integer
Dim rc As Boolean
Dim strLine As String
    rc = True
    With dbTabCode
        .Filter = ""
        If .EOF Then
            rc = False
        Else
            .MoveFirst
            t = 0
            For i = 1 To .RecordCount
            
                strLine = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
                If strLine = "" Then GoTo cont
                If t > 0 Then
                    If GetIndexArStrOneDim(LineList(), strLine) = -1 Then
Aggiungi:
                        
                       ReDim Preserve LineList(t)
                       LineList(t) = strLine
                       t = t + 1
            
                       
                    End If
                Else
                    GoTo Aggiungi
                End If
cont:
                .MoveNext
            Next
        
           
        End If
    
    End With
    GetAllLine = rc
End Function


Public Function GetIndexArStrOneDim(AR() As String, ByVal Val As String) As Long
    Dim i As Long, ei As Long
    
    GetIndexArStrOneDim = -1
 
    On Error Resume Next
        ei = UBound(AR)
        If err.NUMBER <> 0 Then Exit Function
     On Error GoTo 0
 
    For i = 0 To ei
        If UCase(AR(i)) = UCase(Val) Then GetIndexArStrOneDim = i: Exit For
    Next
 
End Function




Public Function SplitTextString(ByVal strObbligatoria As String, ByVal sString As String, ByRef DataType As Variant, ByRef Quanti As Integer) As Boolean
    Dim rc As Boolean
    Dim miaStringa As String
    Dim vettore As Variant
    Dim i As Integer, Somma As Long
    rc = True
    If sString = "" Then GoTo ERR_END
    If Asc(Right(sString, 1)) = 13 Then
        miaStringa = Left(Trim(sString), Len(Trim(sString)) - 1)
    Else
        miaStringa = (Trim(sString))
    End If
    
    vettore = Split(miaStringa, strSeparator)
    
    Somma = 0
    Quanti = 0
    For i = LBound(vettore) To UBound(vettore)
  
            Debug.Print vettore(i) & "  " & i
            
            If InStr(vettore(i), strObbligatoria) Then
                Quanti = Quanti + 1
                ReDim Preserve DataType(Quanti)
                DataType(Quanti - 1) = Trim(vettore(i))
            Else
            End If
    Next
    
    If Quanti = 0 Then
        rc = False
        GoTo ERR_END
    End If
    
    '----------------------------------------------
    ' SCP code
    '----------------------------------------------
    '----------------------------------------------
ERR_END:
    On Error GoTo 0
    SplitTextString = rc
    Exit Function
ERR_CHECK:
    rc = False
    Resume Next
    
End Function




Public Function StringToType(ByVal sString As String, ByRef myDataType As DataType) As Boolean
    Dim rc As Boolean
    Dim miaStringa As String
    Dim vettore As Variant
    Dim i As Integer, Somma As Long, Quanti As Integer
    rc = True
    If sString = "" Then GoTo ERR_END
    If Asc(Right(sString, 1)) = 13 Then
        miaStringa = Left(Trim(sString), Len(Trim(sString)) - 1)
    Else
        miaStringa = (Trim(sString))
    End If
    
    vettore = Split(miaStringa, strSeparator)
    
    Somma = 0
    Quanti = 0
    For i = LBound(vettore) To UBound(vettore)

            Quanti = Quanti + 1
            Debug.Print vettore(i) & "  " & i

        
    Next
    
    If Quanti = 0 Then
        rc = False
        GoTo ERR_END
    End If
    
    '----------------------------------------------
    ' SCP code
    '----------------------------------------------
    
   ' myDataType.Operator = vettore(0)
   ' myDataType.Lot = vettore(1)
   ' myDataType.Code = vettore(2)
   ' myDataType.Inspection = vettore(3)
   ' myDataType.WeightNumber = vettore(4)
   ' myDataType.Weight = vettore(5)
   ' myDataType.Data = Now()
    

    '----------------------------------------------
ERR_END:
    On Error GoTo 0
    StringToType = rc
    Exit Function
ERR_CHECK:
    rc = False
    Resume Next
    
End Function

