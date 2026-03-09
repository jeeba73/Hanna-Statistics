Attribute VB_Name = "Database_ExportHannaCodeTodbCodeFinal"
Option Explicit
Private dbCodePath As String
Private dbCodeName As String


Public Function ExportHannaCodeTodbCodeFinal() As Boolean


        dbCodeName = "dbCode.mdb"
        
        If F_EXPORTHANNACODE.DoShow(dbCodePath, dbCodeName) Then
                
                PopupMessage 2, "Hanna Code Archive exported successfully..."
         
                'SaveSetting App.Title, "ARCHIVIO", "PATH", MyPath
                'SaveSetting App.Title, "ARCHIVIO", "NOME", MydbName
        Else
             
             
        End If






End Function
