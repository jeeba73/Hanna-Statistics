Attribute VB_Name = "Mod_02_MotherSolution"
Option Explicit

        '.MotherSol.bClosed = False
        '.MotherSol.Code = .MRCode
        '.MotherSol.DataPrep = txFormulation(1)
        '.MotherSol.HourPrep = txFormulation(0)
        '.MotherSol.WeekPrep = txFormulation(8)
        '.MotherSol.ExpDays = IIf(.HannaCode.MSEXP = "", 10, CInt(.HannaCode.MSEXP))
        '.MotherSol.DataExp = DateAdd("d", .MotherSol.ExpDays, .MotherSol.DataPrep)
        '.MotherSol.HannaCode = .HannaCode.Code
        '.MotherSol.MSType = .MSType
        '.MotherSol.Note = strNote
        '.MotherSol.Operator = .Operator
        '.MotherSol.QtyLeft = txFormulation(12)
        '.MotherSol.QtyProduced = txFormulation(12)
        '.MotherSol.QtyUsed = 0
        '.MotherSol.PreparationID = PreparationID
        '.MotherSol.Unit = "mL"
    
        
Public Function SaveMotherSolutionInDatabase(ByRef MotherSol As MotherSolution, Optional ByVal ID As Long) As Boolean

Dim rc As Boolean
Dim sString As String

On Error GoTo ERR_SAVE
    
    With dbTabMotherSolution
        If ID > 0 Then
             sString = "ID='" & ID & "'"
        Else
            
            If MotherSol.PreparationID > 0 Then
                sString = "PreparationID='" & MotherSol.PreparationID & "'"
            Else
                sString = "MRCode='" & MotherSol.Code & "' and DataPrep='" & MotherSol.DataPrep & "' and HourPrep='" & MotherSol.HourPrep & "' and DataMS='" & MotherSol.DataMS & "'"
            
            End If
        End If
        
        .filter = ""
        .filter = sString
        
        If .EOF Then
            .AddNew
        End If
        
           !bClosed = MotherSol.bClosed
           !MsType = MotherSol.MsType
           !HannaCode = MotherSol.HannaCode
           !MRCode = MotherSol.Code
           !DataPrep = MotherSol.DataPrep
           !HourPrep = MotherSol.HourPrep
           !PrepWeek = MotherSol.WeekPrep
           !Operator = MotherSol.Operator
           !QtyLeft = MotherSol.QtyLeft
           !QtyProduced = MotherSol.QtyProduced
           !Unit = MotherSol.Unit
           !PreparationID = MotherSol.PreparationID
           !Note = MotherSol.Note
           !DataExp = MotherSol.DataExp
           !DataMS = MotherSol.DataMS
           
           !Bottle = MotherSol.Bottle.EntryBottle
           !BottleLot = MotherSol.Bottle.Lot
           !BottleQty = MotherSol.Bottle.StockQTY
           !BottleID = MotherSol.Bottle.ID

        
            .Update
    
    
    End With
    


ERR_END:
    
   On Error GoTo 0
   SaveMotherSolutionInDatabase = rc
    Exit Function
ERR_SAVE:
    rc = False
    MsgBox Err.Description
    Resume Next

End Function
        
Public Function GetMotherSolutionFromDatabase(ByRef MotherSol As MotherSolution, ByVal ID As Long) As Boolean

Dim rc As Boolean
Dim sString As String
Dim DataExp As Date

On Error GoTo ERR_SAVE

rc = True
    
    With dbTabMotherSolution

        .filter = ""
        .filter = "ID='" & ID & "'"
        
        If .EOF Then

        End If
        
           MotherSol.bClosed = !bClosed
           MotherSol.MsType = IIf(IsNull(Trim(!MsType)), "", Trim(!MsType))
           MotherSol.HannaCode = IIf(IsNull(Trim(!HannaCode)), "", Trim(!HannaCode))
           MotherSol.Code = IIf(IsNull(Trim(!MRCode)), "", Trim(!MRCode))
           MotherSol.DataPrep = IIf(IsNull(Trim(!DataPrep)), "", Trim(!DataPrep))
           MotherSol.HourPrep = IIf(IsNull(Trim(!HourPrep)), "", Trim(!HourPrep))
           MotherSol.WeekPrep = IIf(IsNull(Trim(!PrepWeek)), "", Trim(!PrepWeek))
           MotherSol.Operator = IIf(IsNull(Trim(!Operator)), "", Trim(!Operator))
           MotherSol.QtyLeft = IIf(IsNull(Trim(!QtyLeft)), "", Trim(!QtyLeft))
           MotherSol.Unit = IIf(IsNull(Trim(!Unit)), "", Trim(!Unit))
           MotherSol.QtyProduced = IIf(IsNull(Trim(!QtyProduced)), "", Trim(!QtyProduced))
           MotherSol.PreparationID = IIf(IsNull(Trim(!PreparationID)), "", Trim(!PreparationID))
           MotherSol.Note = IIf(IsNull(Trim(!Note)), "", Trim(!Note))
           MotherSol.DataExp = IIf(IsNull(Trim(!DataExp)), "", Trim(!DataExp))
           MotherSol.DataMS = IIf(IsNull(Trim(!DataMS)), "", Trim(!DataMS))
           MotherSol.MsType = IIf(IsNull(Trim(!MsType)), "", Trim(!MsType))
           
           MotherSol.ID = !ID
           
           MotherSol.Bottle.EntryBottle = IIf(IsNull(Trim(!Bottle)), "", Trim(!Bottle))
           MotherSol.Bottle.Lot = IIf(IsNull(Trim(!BottleLot)), "", Trim(!BottleLot))
           MotherSol.Bottle.StockQTY = IIf(IsNull(Trim(!BottleQty)), "", Trim(!BottleQty))
           MotherSol.Bottle.ID = IIf(IsNull(Trim(!BottleID)), "", Trim(!BottleID))
           
           If IsDate(!DataExp) Then
               DataExp = !DataExp
               If DataExp < FormatDateTime(Now(), 2) Then
                    ' scaduta!
                    !bClosed = True
                     MotherSol.bClosed = True
                End If
            End If
        
            .Update
    
    
    End With
    


ERR_END:
    
   On Error GoTo 0
   GetMotherSolutionFromDatabase = rc
    Exit Function
ERR_SAVE:
    rc = False
    MsgBox Err.Description
    Resume Next

End Function



Public Function FillMotherSolutionTable(ByRef Grid As Grid, ByVal MRCode As String, Optional ByVal MSVolume As Double)

Dim ActualDate As Date
Dim QtyLeft As Double


Dim i As Integer


    ActualDate = FormatDateTime(Now(), 2)
    
    
With Grid
    .Rows = 1
    .AutoRedraw = False
    

    With dbTabMotherSolution
    
        .filter = ""
        .filter = "MRCode='" & MRCode & "' and bClosed='false'"
        
        If .EOF Then
        Else
            .MoveFirst
     
            For i = 1 To .RecordCount
            
            
            
            If Trim(!DataExp) < ActualDate Then GoTo cont:
            
            QtyLeft = CDbl(IIf(IsNull(Trim(!QtyLeft)), 0, Trim(!QtyLeft)))
            
            'If MSVolume > QtyLeft Then GoTo cont:
            
            Grid.AddItem "", False
            Grid.Cell(Grid.Rows - 1, 1).Text = IIf(IsNull(Trim(!DataMS)), "", FormatDataLAT(Trim(!DataMS)))
            Grid.Cell(Grid.Rows - 1, 2).Text = IIf(IsNull(Trim(!QtyProduced)), "", Trim(!QtyProduced) & " mL")
            Grid.Cell(Grid.Rows - 1, 3).Text = IIf(IsNull(Trim(!DataExp)), "", FormatDataLAT(Trim(!DataExp)))
            Grid.Cell(Grid.Rows - 1, 4).Text = IIf(IsNull(Trim(!MRCode)), "", Trim(!MRCode))
            Grid.Cell(Grid.Rows - 1, 5).Text = IIf(IsNull(Trim(!Operator)), "", Trim(!Operator))
            Grid.Cell(Grid.Rows - 1, 6).Text = IIf(IsNull(Trim(!QtyLeft)), "", Trim(!QtyLeft) & " mL")
            Grid.Cell(Grid.Rows - 1, 7).Text = IIf(IsNull(Trim(!Note)), "", Trim(!Note))
            Grid.Cell(Grid.Rows - 1, 8).Text = IIf(IsNull(Trim(!Bottle)), "", Trim(!Bottle))
            Grid.Cell(Grid.Rows - 1, 9).Text = IIf(IsNull(Trim(!BottleLot)), "", Trim(!BottleLot))
            Grid.Cell(Grid.Rows - 1, 10).Text = IIf(IsNull(Trim(!BottleQty)), "", Trim(!BottleQty) & " mL")
            Grid.Cell(Grid.Rows - 1, 11).Text = !ID
            
            '.Cell(0, 1).Text = "DataMS"
            '.Cell(0, 2).Text = "QtyProduced"
            '.Cell(0, 3).Text = "DataExp"
            '.Cell(0, 4).Text = "MRCode"
            '.Cell(0, 5).Text = "Operator"
            '.Cell(0, 6).Text = "QtyLeft"
            '.Cell(0, 7).Text = "Note"
            '.Cell(0, 8).Text = "Bottle Number"
            '.Cell(0, 9).Text = "BottleLot"
            '.Cell(0, 10).Text = "Bottle Qty"
            '.Cell(0, 11).Text = "ID"
            
            Grid.Cell(Grid.Rows - 1, 2).BackColor = vbColorResults
            Grid.Cell(Grid.Rows - 1, 6).BackColor = vbColorResults
    
cont:
                .MoveNext
            Next
        End If
    End With
    For i = 1 To .Cols - 1
        .Column(i).AutoFit
        .Column(i).Width = .Column(i).Width * 1.5
    Next
    .Column(11).Width = 0
    .Refresh
    .AutoRedraw = True
    
End With
End Function

