Attribute VB_Name = "QRCodeTypeReadings"
Option Explicit
Private SettingName As String



Public Function PrintStandardLabel(ByRef MyQRCode As Qr01, ByVal bManual As Boolean)
Dim PREP_PATH As String
Dim HannaCode As String
Dim MaxHannaCode As String
Dim i As Integer
Dim rc As Boolean
Dim stringNote As String
Dim m_rc As Boolean
Dim RfPfileName As String


    CloseSettingDataFile
    
    rc = False
    
    
    If MyOperatore.Name = "" Then

    If frmLogin.DoShow Then
            
    Else
        Exit Function
    End If

    End If

    MyQRCode.Operator = MyOperatore.Name
    stringNote = MyQRCode.Note
    If F_InputBox.DoShow("Enter QRCode Note", "QRCode : " & MyQRCode.Code, , , , stringNote) Then MyQRCode.Note = stringNote
     

    rc = DoPrintStandardLabel(MyQRCode, False, bManual)

        
    If rc Then
        PopupMessage 2, "Printing Label....", MyQRCode.Code
    End If

    CloseSettingDataFile
    
End Function
