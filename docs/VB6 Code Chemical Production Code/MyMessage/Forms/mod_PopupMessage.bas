Attribute VB_Name = "mod_PopupMessage"
Option Explicit

Public UploadDownloadMessageCounter As Integer
Public MessageInfoTime As Integer
Public myUploadMessageForm() As Form
Public myDownloadMessageForm() As Form
Public myInfoMessage() As Form
Public MessageTime   As Integer
Public Const FadeTime = 25
Public Sub PopupMessage(ByVal Index As Integer, ByVal sMessage As String, Optional ByVal sStazione As String, Optional ByVal bRed As Boolean, Optional ByVal sTitle As String, Optional ByVal MyImage As Image, Optional ByVal bButton As Boolean)

On Error GoTo ERR_POP

    
    UploadDownloadMessageCounter = UploadDownloadMessageCounter + 1
    
    Select Case Index
        Case 0 ' upload
        Case 1 ' download
        Case 2
            ReDim Preserve myInfoMessage(1 To UploadDownloadMessageCounter) As Form
            Set myInfoMessage(UploadDownloadMessageCounter) = New MessageInfo
            myInfoMessage(UploadDownloadMessageCounter).lDescription = sMessage
            If sTitle <> "" Then myInfoMessage(UploadDownloadMessageCounter).lTitle = sTitle
            myInfoMessage(UploadDownloadMessageCounter).DoShow bRed, MyImage, bButton 'SHOW MESSAGE

    End Select
    

ERR_END:
    On Error GoTo 0
    MessageInfoTime = 1700
    Exit Sub
ERR_POP:
    MsgBox err.Description
    Resume Next

End Sub

