Attribute VB_Name = "mod_Excel_DeleteAllTabs"
Option Explicit


Public Function DeleteAllTabCode()


    dbCode.Execute "DELETE * FROM TabCode"
    DoEvents
    dbTabCode.Close
    dbTabCode.Open "SELECT *  FROM TabCode ORDER BY Code", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    
    DoEvents
   
End Function

Public Function DeleteAllTabMR()


    dbCode.Execute "DELETE * FROM TabMR"
    DoEvents
    dbTabCode.Close
    dbTabCode.Open "SELECT *  FROM TabMR ORDER BY Code", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    
  
    DoEvents
   
End Function

Public Function DeleteAllTabFrasiH()


    dbCode.Execute "DELETE * FROM TabFrasiH"
    DoEvents
    dbTabFrasiH.Close
    dbTabFrasiH.Open "SELECT *  FROM TabFrasiH ORDER BY Code", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    
    DoEvents
   
End Function



