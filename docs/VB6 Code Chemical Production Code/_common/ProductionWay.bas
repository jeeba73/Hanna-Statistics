Attribute VB_Name = "ProductionWay"
Option Explicit

Public Function SetProductionWay(ByVal Line As String, ByRef uProductionWay() As ProdWay) As Boolean

Dim rc As Boolean
Dim i As Integer
Dim Count As Integer
    Line = Trim(Line)
    rc = False
    ReDim uProductionWay(0)
    With dbTabMachine
        .filter = ""
        .filter = "Line='" & Line & "'"
        
       
        
        If .EOF Then
            MessageInfoTime = 2000
            PopupMessage 2, "Line : " & Line & " doesn't have a Machine List!!", , True, "Machine"
            
        Else
tutte:
            .MoveFirst
            
               
                
                Count = 1
                For i = 1 To .RecordCount
                    If InStr(Trim(!Line), Line) Then
                        ReDim Preserve uProductionWay(Count)
                        uProductionWay(Count).Line = Line
                        uProductionWay(Count).Production = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
                        uProductionWay(Count).Head = IIf(IsNull(Trim(!HEADS_NUMBER)), 0, Trim(!HEADS_NUMBER))
                       'uProductionWay(Count).Speed = IIf(IsNull(Trim(!Speed)), 0, Trim(!Speed))
                        rc = True
                        Count = Count + 1
                    End If
                    .MoveNext
                Next
    
        End If
    End With
    SetProductionWay = rc

End Function


Public Function GetSpeedbyLine(ByVal ProductionWay As String, ByVal Line As String) As String

Dim i As Integer
    Line = Trim(Line)

    GetSpeedbyLine = 0

    With dbTabProductionWay
        .filter = ""
        .filter = "Line='" & Line & "' and ProductionWay='" & ProductionWay & "'"
        If .EOF Then
        Else
            .MoveFirst
            For i = 1 To .RecordCount
                If InStr(Trim(!Line), Line) Then
                    GetSpeedbyLine = IIf(IsNull(!Speed), "0", !Speed)
                    Exit For
                End If
                .MoveNext
            Next
        End If
    End With
End Function
Public Function SetComboMachine(ByVal Line As String, ByRef Combo As ComboBox) As Boolean

Dim rc As Boolean
Dim i As Integer
Dim Count As Integer
    Line = Trim(Line)
    rc = False
    
    Combo.Clear
    
    With dbTabProductionWay
        .filter = ""
        .filter = ""
        If .EOF Then
            MessageInfoTime = 2500
            PopupMessage 2, "Please fill Production Way Database first!", , True
            SetComboMachine = True
            Exit Function
        End If
        
        If .EOF Then
            MessageInfoTime = 2000
            PopupMessage 2, "Line : " & Line & " doesn't have a production Way!!", , True, "Packaging"
            
        Else
tutte:
            Combo.AddItem "Select Machine"
            .MoveFirst
            
               
                
                Count = 1
                For i = 1 To .RecordCount
                    If InStr(Trim(!Line), Line) Then
                        
                       Combo.AddItem IIf(IsNull(Trim(!ProductionWay)), "", Trim(!ProductionWay))
                       
                        rc = True
                        Count = Count + 1
                    End If
                    .MoveNext
                Next
    
        End If
    End With
    SetComboMachine = rc

End Function
