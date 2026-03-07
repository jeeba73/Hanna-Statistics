Attribute VB_Name = "QRCodeReaderQC"
Option Explicit

Public Type QRCodeType
    
    Recipe          As String
    code            As String
    Lot             As String
    Exp             As String
    Operator        As String
    date            As String
    Time            As String
    
    QC              As String
    Note            As String
    Qty             As String
    BoxQty          As String
    BoxNr           As String
    
    
    Tablet          As String
    
    STDv            As String
    STDn            As String
    STD_ID              As String
    
    
    FileName        As String
    
    bCheck           As Boolean
    
End Type

Public iQRCodeType As QRCodeType
Public iQRCodeTypeClean As QRCodeType







Public Function GetSTDQRCodeFromString(ByVal sString As String, ByRef myQrDataType As QRCodeType) As Boolean
Dim rc As Boolean
rc = True
On Error GoTo ERR_CHECK:


    If InStr(sString, sQRSeparator) Then
        Call STDQRCodeString(sString, myQrDataType)
    Else
        rc = False
    End If

    
ERR_END:
    On Error GoTo 0
    GetSTDQRCodeFromString = rc
    Exit Function
ERR_CHECK:
    rc = False
    Resume Next
End Function

Public Function GetQRCodeFromString(ByVal sString As String, ByRef myQrDataType As QRCodeType) As Boolean
Dim rc As Boolean
rc = True
On Error GoTo ERR_CHECK:


    If InStr(sString, sQRSeparator) Then
        Call QRCodeString(sString, myQrDataType)
    Else
        rc = False
    End If

    
ERR_END:
    On Error GoTo 0
    GetQRCodeFromString = rc
    Exit Function
ERR_CHECK:
    rc = False
    Resume Next
End Function

Private Function QRCodeString(ByVal sString As String, myQrDataType As QRCodeType) As Boolean
    Dim rc As Boolean
    Dim miaStringa As String
    Dim vettore As Variant
    Dim LastVettore As Variant
    Dim i As Integer, Somma As Long, Quanti As Integer, QuantiLastVettore As Integer
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
        
        
            'Recipe          As String
            'Code            As String
            'Lot             As String
            'Exp             As String
            'Operator        As String
            'Date            As String
            'Time            As String
            
            'QC              As String
            'Note            As String
            'Qty             As String
            'BoxQty          As String
            'BoxNr           As String
                

            .Recipe = vettore(0)
            .code = vettore(1)
            .Qty = vettore(2)
            .Lot = vettore(3)
            .Exp = vettore(4)
            .Operator = vettore(5)
            .date = vettore(6)
            .Time = vettore(7)
             If UBound(vettore) > 7 Then
                .QC = vettore(8)
               
             End If
           
            .bCheck = True
        End With
    Else
        rc = False
        GoTo ERR_END
    End If

    '----------------------------------------------
ERR_END:
    On Error GoTo 0
    If rc Then
        PopupMessage 2, "BarCode Reading....", , , myQrDataType.code & " | " & myQrDataType.Lot
    End If
    
    QRCodeString = rc
    Exit Function
ERR_CHECK:
    rc = False
    
    myQrDataType.bCheck = False
    
    Resume Next
    
End Function

Private Function STDQRCodeString(ByVal sString As String, myQrDataType As QRCodeType) As Boolean
    Dim rc As Boolean
    Dim miaStringa As String
    Dim vettore As Variant
    Dim LastVettore As Variant
    Dim i As Integer, Somma As Long, Quanti As Integer, QuantiLastVettore As Integer
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
    If UBound(vettore) >= 5 Then
        With myQrDataType
        
        
            'Recipe          As String
            'Code            As String
            'Lot             As String
            'Exp             As String
            'Operator        As String
            'Date            As String
            'Time            As String
            
            'QC              As String
            'Note            As String
            'Qty             As String
            'BoxQty          As String
            'BoxNr           As String
                

            .code = vettore(0)
            .STDn = vettore(1)
            .STDv = vettore(2)
            .STD_ID = vettore(3)
            .Operator = vettore(4)
           
            .bCheck = True
        End With
    Else
        rc = False
        GoTo ERR_END
    End If

    '----------------------------------------------
ERR_END:
    On Error GoTo 0
    If rc Then
        PopupMessage 2, "BarCode Reading....", , , myQrDataType.code & " | STD" & myQrDataType.STDn
    End If
    
    STDQRCodeString = rc
    Exit Function
ERR_CHECK:
    rc = False
    
    myQrDataType.bCheck = False
    
    Resume Next
    
End Function




