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

        Case "TabFGrevisionHistory"
            '-------------------------------------------------------
               strSQL = "CREATE TABLE TabFGRevisionHistory (" & _
                "ID INT IDENTITY PRIMARY KEY NOT NULL , RevDate  DateTime, Code varchar(40) WITH COMPRESSION, RevNumber varchar(30) WITH COMPRESSION, Type varchar(100) WITH COMPRESSION, Description varchar(255) WITH COMPRESSION ,Operator varchar(100) WITH COMPRESSION   )"
            '-------------------------------------------------------------
        
            dbCode.Execute strSQL
            dbCode.Close
            GoTo ERR_END:
        
        Case "TabFinishGood"
            '-------------------------------------------------------
               strSQL = "CREATE TABLE TabFinishGood (" & _
                "ID INT IDENTITY PRIMARY KEY NOT NULL ,  Code varchar(100) WITH COMPRESSION," & _
                "Description varchar(100) WITH COMPRESSION , Method varchar(255) WITH COMPRESSION," & _
                "Rangeppm varchar(20) WITH COMPRESSION , RefMeter varchar(100) WITH COMPRESSION,RefMeterDescription varchar(255) WITH COMPRESSION," & _
                "RefSTD varchar(100) WITH COMPRESSION , Wavelength varchar(20) WITH COMPRESSION," & _
                "Cell varchar(100) WITH COMPRESSION , RefSTDNote varchar(255) WITH COMPRESSION,RefSTDNote2 varchar(255) WITH COMPRESSION," & _
                "gdl varchar(20) WITH COMPRESSION , Slope varchar(20) WITH COMPRESSION," & _
                "OrdinateIntersect varchar(20) WITH COMPRESSION , ReagentBlank varchar(20) WITH COMPRESSION," & _
                "MethVar varchar(20) WITH COMPRESSION , ConfInt varchar(20) WITH COMPRESSION,StdDeviation varchar(20) WITH COMPRESSION," & _
                "LastEdit  DateTime)"
            '-------------------------------------------------------------
            
            dbCode.Execute strSQL
            dbCode.Close
            GoTo ERR_END:
        
        
        
        
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
            
            
            
            
            
            
            
   End Select

    dbChemicalQC.Execute strSQL
    
    dbChemicalQC.Close
               
ERR_END:
    On Error GoTo 0
    AddTable = rc
    Exit Function
ERR_ADD:
    
    MsgBox Err.Description
    rc = False
    Resume ERR_END
End Function
