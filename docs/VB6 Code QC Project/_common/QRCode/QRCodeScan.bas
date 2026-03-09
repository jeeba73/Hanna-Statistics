Attribute VB_Name = "QRCodeScan"
Option Explicit



'--------------------------------------------------------------------
' check scanned QRCode
'---------------------------------------------------------------------


Public Function QRCodeToTabReport(ByRef QRCode As QRCodeType, Optional ByVal bClosed As Boolean) As Boolean

Dim rc As Boolean
Dim i As Integer
Dim sString As String

On Error GoTo ERR_GET:

        rc = True
        
        With QRCode

            sString = "Code ='" & .Code & "' and Lot ='" & .Lot & "' and Exp ='" & .Exp & "'"
    
        End With
        
        With dbTabReport
            .filter = ""
            .filter = sString
            If .EOF Then
                rc = False
            Else
                QRCode.fileName = IIf(IsNull(Trim(!Nomefile)), "", Trim(!Nomefile))
            End If
         
         End With
         
        
ERR_END:
    On Error GoTo 0

    QRCodeToTabReport = rc
    Exit Function
ERR_GET:
    rc = False
    Resume Next
End Function

Public Function SearchQCInTab(ByRef UserQrCode As QRCodeType, ByRef Grid As Grid) As Boolean
    Dim rc As Boolean
    Dim i As Integer
    On Error GoTo ERR_QR:
    
    rc = True
    
    With Grid
        If .Rows > 1 Then
            
            For i = 1 To .Rows - 1
                If .Cell(i, 2).Text = UserQrCode.Code And .Cell(i, 16).Text = UserQrCode.Lot Then
                    .Cell(i, 2).EnsureVisible
                    .Cell(i, 2).SetFocus
                    Exit For
                End If
            Next
            rc = False
        End If
    End With
ERR_END:
    On Error GoTo 0

    SearchQCInTab = rc
    Exit Function
ERR_QR:
    rc = False
    Resume ERR_END:


End Function


