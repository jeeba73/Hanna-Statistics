Attribute VB_Name = "BarcodeReader"
Option Explicit

Public sQRSeparator As String

'Private Type Barcode
'    Code                    As String
'    Description             As String
'    Manufacturer            As String
'    ManufacturerCode        As String
'    ManufacturerLot         As String
'    DeliveryDate            As String  '
'    Package                 As String
'    WeekDelPackageNumber    As String
'End Type


Public myQrDataType As Barcode


Public Function GetSeparator()
    
        sQRSeparator = GetSetting(App.Title, "QRCODE", "Separatore", ":")
        bOpenProductClassificationAfterScan = GetSetting(App.Title, "BarcodeReader", "bOpenProductClassificationAfterScan", False)

End Function



Public Function TestReader()

Dim sString As String
MessageInfoTime = 3000
PopupMessage 2, "Please set your barcode Reader in Keyboard Wedge mode..."
DoEvents
If F_InputBox.DoShow("Scan your Code...", "BarCode Reader", , , , sString) Then
    MessageInfoTime = 4500
    PopupMessage 2, "Reading : " & sString, , , "BarCode Reader"
End If


End Function


Public Function SetQRSeparator() As Boolean
Dim rc As Boolean
Dim sString As String


rc = True

sString = sQRSeparator

If F_InputBox.DoShow("Please Set Separator Char", "BarCode Reader", , , , sString) Then
    If Len(sString) > 1 Then
err:
        MessageInfoTime = 4500
        PopupMessage 2, "Warning : String must be 1 char...", "QRCode"
        rc = False
        
    ElseIf Len(sString) = 1 Then
    
        SaveSetting App.Title, "QRCODE", "Separatore", sString
        sQRSeparator = sString
        MessageInfoTime = 2500
        PopupMessage 2, "Barcode Settings correctly done..." & vbCrLf & "QRCode Separator = " & sQRSeparator, , , "QRCode"
        rc = True
    Else
        GoTo err:
    End If
End If

SetQRSeparator = rc
End Function



Public Function CheckQRCode_HANNA(ByVal sString As String, ByRef UserQrCode As Barcode) As Boolean
Dim rc As Boolean
rc = True
On Error GoTo ERR_CHECK:


    If InStr(sString, sQRSeparator) Then
        Call QRCodeString(sString, myQrDataType)
    Else
        rc = False
    End If

    UserQrCode = myQrDataType
    
    
ERR_END:
    On Error GoTo 0
    CheckQRCode_HANNA = rc
    Exit Function
ERR_CHECK:
    rc = False
    Resume Next
End Function

Private Function QRCodeString(ByVal sString As String, myQrDataType As Barcode) As Boolean
    Dim rc As Boolean
    Dim miaStringa As String
    Dim vettore As Variant
    Dim LastVettore As Variant
    Dim i As Integer, Somma As Long, Quanti As Integer, QuantiLastVettore As Integer
    
    On Error GoTo ERR_CHECK
    rc = True
    If sString = "" Then GoTo ERR_END
    If Asc(Right(sString, 1)) = 13 Then
        miaStringa = Left(Trim(sString), Len(Trim(sString)) - 1)
    Else
        miaStringa = (Trim(sString))
    End If
    
    vettore = Split(miaStringa, sQRSeparator)
    
    Somma = 0
    Quanti = 0
    For i = LBound(vettore) To UBound(vettore)

            Quanti = Quanti + 1
    Next
    
    LastVettore = Split(vettore(Quanti - 1), " ")
    QuantiLastVettore = UBound(LastVettore)
    If Quanti = 0 Then
        rc = False
        GoTo ERR_END
    End If
    
    '----------------------------------------------
    ' Hanna QR code
    '----------------------------------------------
    If UBound(vettore) >= 6 Then
        With myQrDataType
            .Code = vettore(0)
            .ChemicalName = vettore(1)
            .Manufacturer = vettore(2)
            .ManufacturerCode = vettore(3)
            .ManufacturerLot = vettore(4)
            .DeliveryDate = vettore(5)
            .QtyDelivered = vettore(6)
            If QuantiLastVettore > 0 Then
                .Package = LastVettore(QuantiLastVettore)
                .WeekDelPackageNumber = LastVettore(0)
            End If
            
        End With
    Else
        rc = False
        GoTo ERR_END
    End If

    '----------------------------------------------
ERR_END:
    On Error GoTo 0
    QRCodeString = rc
    Exit Function
ERR_CHECK:
    rc = False
    MsgBox err.Description
    Resume Next
    
End Function

