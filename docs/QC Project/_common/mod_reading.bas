Attribute VB_Name = "mod_reading"
Option Explicit




Public Function CheckRows(ByRef Rows As Long, ByVal SettingName As String, ByVal PATH As String) As Boolean

    
    Dim i As Integer
    Dim t As Integer
    Dim sSTD As String
    Dim sString As String
    
    ' Readings : check standard in table....
    
    If SettingName = "" Then Exit Function
    CloseSettingDataFile
    
    t = 1
    i = 0
    Do
        
        sSTD = GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col1", "", PATH)
        sString = GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col2", "", PATH)
        i = i + 1
        
        If i >= 500 Then Exit Function
        
    Loop Until sSTD = "" And sString = ""
    
    Rows = i - 1
    
    CloseSettingDataFile

    If Rows > 1 Then
        
         SaveSettingData SettingName, "Reading QC", "Grd2 Rows", Rows, PATH
    
    End If
    
    CloseSettingDataFile

End Function


