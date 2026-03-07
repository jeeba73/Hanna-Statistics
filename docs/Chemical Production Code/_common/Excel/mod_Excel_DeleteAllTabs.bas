Attribute VB_Name = "mod_Excel_DeleteAllTabs"
Option Explicit

Public Function DeleteAllTabFrasiH()


    dbCode.Execute "DELETE * FROM TabFrasiH"
    DoEvents
    dbTabFrasiH.Close
    dbTabFrasiH.Open "SELECT *  FROM TabFrasiH ORDER BY Code", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    
    DoEvents
   
End Function

Public Function DeleteAllTabCode()


    dbCode.Execute "DELETE * FROM TabCode"
    DoEvents
    dbTabCode.Close
    dbTabCode.Open "SELECT *  FROM TabCode ORDER BY Code", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    
    DoEvents
   
End Function
Public Function DeleteAllTabRecipe()


    dbCode.Execute "DELETE * FROM TabRecipe"
    DoEvents
    dbTabRecipe.Close
    dbTabRecipe.Open "SELECT *  FROM TabRecipe order by Code ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
   
    DoEvents
   
End Function
