Attribute VB_Name = "Locations"
Option Explicit

Public Function AddLocationInCombo(ByRef cmb As ComboBox) As Boolean

Dim i As Integer

cmb.Clear

With dbTabLocation
    .Close
    .Open "SELECT *  FROM TabLocation order by Code ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    .Filter = ""
    .Filter = ""
    
    If .EOF Then
        
    Else
        .MoveFirst
        For i = 1 To .RecordCount
        
            If IsNull(Trim(!Code)) Or Trim(!Code) = "" Then
            Else
                cmb.AddItem Trim(!Code)
            End If
            
            .MoveNext
        Next
    
        cmb.ListIndex = 0
    End If

End With


End Function





Public Function AddLocationInDatabase(ByRef Code As String) As Boolean

Dim rc As Boolean

Code = Trim(Code)

If Code = "" Then Exit Function

rc = True
With dbTabLocation
    .Close
    .Open "SELECT *  FROM TabLocation order by Code ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    .Filter = ""
    .Filter = "Code='" & Replace(Code, "'", "''") & "'"
    
    If .EOF Then
        If F_MsgBox.DoShow("Add new Location in Database?", Code) Then
            .AddNew
            !Code = Code
            .Update
        End If
    Else
        rc = False
    End If

End With
AddLocationInDatabase = rc


End Function


Public Function DeleteLocationFromDatabase(ByRef Code As String) As Boolean

Dim rc As Boolean

Code = Trim(Code)

rc = True

With dbTabLocation
    .Close
    .Open "SELECT *  FROM TabLocation order by Code ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    .Filter = ""
    .Filter = "Code='" & Replace(Code, "'", "''") & "'"
    
    If .EOF Then
        rc = False
    Else
        .MoveFirst
       .Delete
       .Update
       
       PopupMessage 2, "Record deleted....", "Location"
    End If

End With
DeleteLocationFromDatabase = rc


End Function








