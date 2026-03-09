Attribute VB_Name = "mod_ADDTABLE"
Public Function AddTable(ByVal NameTable As String) As Boolean
Dim rc As Boolean
Dim strSQL As String

On Error GoTo ERR_ADD
    rc = True
    Select Case NameTable

       ' Case "TabRecipeRevisionHistory"
            '-------------------------------------------------------
            '   strSQL = "CREATE TABLE TabRecipeRevisionHistory (" & _
           '     "ID INT IDENTITY PRIMARY KEY NOT NULL, RevDate  DateTime, Recipe  varchar(100) WITH COMPRESSION, RevNumber  varchar(100) WITH COMPRESSION, RevType  varchar(100) WITH COMPRESSION, Note  varchar(255) WITH COMPRESSION, Operator  varchar(100) WITH COMPRESSION )"
            '-------------------------------------------------------------

        Case "TabRecipeRevisionHistory"
            '-------------------------------------------------------
               strSQL = "CREATE TABLE TabRecipeRevisionHistory (" & _
                "ID INT IDENTITY PRIMARY KEY NOT NULL , RevDate  DateTime, Recipe varchar(40) WITH COMPRESSION, RevNumber varchar(30) WITH COMPRESSION, Type varchar(100) WITH COMPRESSION, Description varchar(255) WITH COMPRESSION ,Operator varchar(100) WITH COMPRESSION   )"
            '-------------------------------------------------------------

        Case "TabPreparationNotes"
            '-------------------------------------------------------
               strSQL = "CREATE TABLE TabPreparationNotes (" & _
                "ID INT IDENTITY PRIMARY KEY NOT NULL , NoteDate  DateTime, FileName varchar(100) WITH COMPRESSION, Type varchar(100) WITH COMPRESSION , Description varchar(255) WITH COMPRESSION ,Operator varchar(100) WITH COMPRESSION   )"
            '-------------------------------------------------------------
            dbChemicalProduction.Execute strSQL
            dbChemicalProduction.Close
            GoTo ERR_END:
        
        
        Case "TabProductionNotes"
            '-------------------------------------------------------
               strSQL = "CREATE TABLE TabProductionNotes (" & _
                "ID INT IDENTITY PRIMARY KEY NOT NULL , NoteDate  DateTime, FileName varchar(100) WITH COMPRESSION, Type varchar(100) WITH COMPRESSION, Description varchar(255) WITH COMPRESSION ,Operator varchar(100) WITH COMPRESSION   )"
            '-------------------------------------------------------------
            dbChemicalProduction.Execute strSQL
            dbChemicalProduction.Close
            GoTo ERR_END:
            
        Case "TabMachine"
            '-------------------------------------------------------
               strSQL = "CREATE TABLE TabMachine (" & _
                "ID INT IDENTITY PRIMARY KEY NOT NULL , MACHINE varchar(50) WITH COMPRESSION, DESCRIPTION varchar(100) WITH COMPRESSION, HEADS_NUMBER varchar(10) WITH COMPRESSION ,Model varchar(100) WITH COMPRESSION ,SerialNumber varchar(100) WITH COMPRESSION ,Line varchar(100) WITH COMPRESSION   )"
            '-------------------------------------------------------------
           ' dbCode.Execute strSQL
           ' dbCode.Close
           ' GoTo ERR_END:
                        
            
   End Select

    dbCode.Execute strSQL
    
    dbCode.Close
               
ERR_END:
    On Error GoTo 0
    AddTable = rc
    Exit Function
ERR_ADD:
    
    'MsgBox err.Description
    rc = False
    Resume ERR_END
End Function
