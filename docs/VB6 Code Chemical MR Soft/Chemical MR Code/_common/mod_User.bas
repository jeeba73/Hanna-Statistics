Attribute VB_Name = "mod_User"
Option Explicit

Public Function CheckPrivilege(ByVal MyPrivilege As Integer) As Boolean
Dim rc As Boolean

rc = True

If MyPrivilege = 2 Then MyPrivilege = 1 ' questo perchè TCO e Laboratory hanno gli stessi privilegi...

If MyOperatore.IndexPrivilege < MyPrivilege Then rc = False
If MyOperatore.Name = "" Then rc = False

If rc = False Then
    ' richiamo la form Login...
    frmLogin.DoShow MyPrivilege
    rc = True
    If MyOperatore.IndexPrivilege < MyPrivilege Then rc = False
    If MyOperatore.Name = "" Then rc = False


End If

If rc = False Then
    MessageInfoTime = 2000
    
    PopupMessage 2, "Warning : minimum authentication required : " & GetPosition(MyPrivilege), , True
    DoEvents
End If


CheckPrivilege = rc

End Function

Public Function GetPosition(ByVal IndexPrivilege As Long) As String

Select Case IndexPrivilege
    Case 0
        GetPosition = "Operator"
    Case 1
        GetPosition = "Production Manager"
    Case 2
        GetPosition = "Line Leader"
    Case 3
        GetPosition = "Administrator"
End Select



End Function
