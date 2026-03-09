Attribute VB_Name = "ProgramSettings"
Option Explicit

Public Function GetProgramSettings()


    ' release 1.2.0
    
    
    dbCodeName = GetSetting(App.Title, "Update", "dbName", "dbCodeQC.mdb")
    'dbCodeName = "dbCodeQC.mdb"
    
    
    
    
    bStampaOk = GetSetting(App.Title, "LABEL PRINTER", "bUtilizzo", False)
    If bStampaOk Then
        MyPathLabel_Brother = GetSetting(App.Title, "PATH", "TEMPLATE LABEL", "")
        bStampanteSelezionata = GetSetting(App.Title, "LABEL PRINTER", "bStampanteSelezionata", False)
        
        SearchInfoLabelPrinter
    End If
    
 
End Function

