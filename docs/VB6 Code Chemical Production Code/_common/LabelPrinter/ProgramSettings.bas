Attribute VB_Name = "ProgramSettings"
Option Explicit

Public Function GetProgramSettings()

    
    bStampaOk = GetSetting(App.Title, "LABEL PRINTER", "bUtilizzo", False)
    If bStampaOk Then
        MyPathLabel_Brother = GetSetting(App.Title, "PATH", "TEMPLATE LABEL", "")
        bStampanteSelezionata = GetSetting(App.Title, "LABEL PRINTER", "bStampanteSelezionata", False)
        
        SearchInfoLabelPrinter
    End If
    
    
    CheckLotNumberType
 
End Function

