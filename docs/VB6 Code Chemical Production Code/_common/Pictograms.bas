Attribute VB_Name = "Pictograms"
Option Explicit

Public Function SetStringAndPictogram(ByRef PictogramCodes() As String, ByRef strClassificationCodes As String, ByRef strPictogramCodes As String, ByRef ClassificationCodes() As Variant, ByRef frm As Form)
Dim Quanti As Integer
Dim rc As Boolean
    rc = False
    If Len(strClassificationCodes) > 0 Then
                        
        Call SplitTextStringClassification("H", strClassificationCodes, ClassificationCodes(), Quanti)
    
        If Quanti > 0 Then
            
            rc = True
            Call SetPictogramCodes(PictogramCodes, ClassificationCodes(), Quanti)
            
            If Quanti > 0 Then Call Setpictogram(PictogramCodes, frm)
        End If
    End If
    If strClassificationCodes = "" And strPictogramCodes = "" Then
        MessageInfoTime = 20
        PopupMessage 2, "No Classification..."
        rc = False
    End If
    
    SetStringAndPictogram = rc
    
End Function

Public Function SetPictogramCodes(ByRef PictogramCodes() As String, ByRef ClassificationCodes() As Variant, ByRef Quanti As Integer)
Dim i As Integer
Dim strPictogram As String
Dim t As Integer
    
    ReDim PictogramCodes(0)
    Quanti = 0
    t = 0

    For i = LBound(ClassificationCodes) To UBound(ClassificationCodes)
    
        With dbTabFrasiH
            .filter = ""
            .filter = "Code='" & ClassificationCodes(i) & "'"
            If .EOF Then
                strPictogram = ""
            Else
                strPictogram = IIf(IsNull(Trim(!Pictogram)), "", Trim(!Pictogram))
                If InStr(strPictogram, "GH") Then
                
                
                    If t > 0 Then
                        If GetIndexArStrOneDim(PictogramCodes(), strPictogram) = -1 Then
Aggiungi:
                            
                           ReDim Preserve PictogramCodes(t)
                           PictogramCodes(t) = strPictogram
                           t = t + 1
                
                           Quanti = t
                        End If
                    Else
                        GoTo Aggiungi
                    End If
                    
                End If
            End If

        End With
    
    Next


End Function


Public Function Setpictogram(ByRef PictogramCodes() As String, ByRef frm As Form)
Dim i As Integer

For i = LBound(PictogramCodes) To UBound(PictogramCodes)
    If PictogramCodes(i) = "" Then Exit Function
    Call LoadPictogram(PictogramCodes(i), i, frm)
Next




End Function


Private Sub LoadPictogram(ByVal Code As String, ByVal Index As Integer, ByRef frm As Form)
Dim strFile As String
    If Code = "" Then Exit Sub
    strFile = PathPictograms & "\" & Code & ".ico"
    frm.Pictogram(Index) = LoadPicture(strFile)
    frm.Pictogram(Index).Visible = True
End Sub
