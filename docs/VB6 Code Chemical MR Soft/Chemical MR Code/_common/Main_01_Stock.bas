Attribute VB_Name = "Main_01_Stock"
Option Explicit

Public Function GetStockFromDatabase(ByRef Grid As Grid, ByVal bClosed As Boolean, Optional ByVal UserMRcode As String, Optional strLotto As String, Optional ByVal bMotherSolution As String, Optional ByVal MRQty As Double, Optional ByVal bTutti As Boolean) As Boolean
Dim i As Integer
Dim t As Integer
Dim x As Integer
Dim Count As Integer
Dim MNP As String
Dim sString As String


Dim iWarehouse() As WareHouseEntry
CloseSettingDataFile



    If UCase(UserMRcode) = "SEARCH" Then UserMRcode = ""
    
    If bClosed = False Then
    
        sString = "bClosed='false'"
    Else
    End If

    If UserMRcode = "" Then
       ' sString = ""
        'Grid.Column(2).Width = 150
    Else
    
    
        sString = sString & " and Code='" & Replace(Trim(UserMRcode), "'", "''") & "'"
        
      
        
    End If

    If strLotto <> "" Then
    
        sString = sString & " and Lot='" & Replace(Trim(strLotto), "'", "''") & "'"
        
    End If
        
    
        
               
        


With Grid

    .Rows = 1
    .AutoRedraw = False
   
    
    With dbTabMRWarehouse
        .Close
        .Open "SELECT *  FROM TabMRWarehouse order by MREXP", dbChemicalMR, adOpenKeyset, adLockOptimistic, adCmdText
        .filter = ""
        .filter = sString
        
        If .EOF Then
            Count = 0
        Else
            Count = .RecordCount
            .MoveLast
        End If
        
        
     '   If bTuttiRecord = False And bClosed Then
         '  Count = IIf(Count > 30, 30, Count)
      '  End If

               
        '.Cell(0, 1).Text = "MR Code"
        '.Cell(0, 2).Text = "Bottle"
        '.Cell(0, 3).Text = "Description"
        '.Cell(0, 4).Text = "MR Lot"
        '.Cell(0, 5).Text = "Purity  %"
        '.Cell(0, 6).Text = "Value"
    
        '.Cell(0, 7).Text = "Unit"

        '.Cell(0, 8).Text = "Location"
        '.Cell(0, 9).Text = "Stock QTY"
        '.Cell(0, 10).Text = "Stock Unit"
        '.Cell(0, 11).Text = "Arrived"
        
        '.Cell(0, 12).Text = "Open"
        '.Cell(0, 13).Text = "Finished"
        '.Cell(0, 14).Text = "Supplier EXP"
        '.Cell(0, 15).Text = "MR EXP"
        '.Cell(0, 16).Text = "Status"
        '.Cell(0, 17).Text = "Note"

        '.Cell(0, 18).Text = "ID"
        
        Dim PurityPerc  As Double
        Dim StockQTY    As Double
        Dim stockUnit   As String
        Dim ChemicalMRCode As String
        Dim MRCode()    As MRType
        
        
        ReDim MRCode(0)
        
        
        For i = 1 To Count
            If IsNull(Trim(!StockQTY)) Or Trim(!StockQTY) = "" Then GoTo cont:
            PurityPerc = 100
            If IsNumeric(IIf(IsNull(Trim(!Purity)), "", Trim(!Purity))) Then
                
                PurityPerc = CheckDot(Trim(!Purity))
                If CheckDot(Trim(!Purity)) < 1 Then
                    PurityPerc = CheckDot(Trim(!Purity)) * 100
                End If
            End If
            
            If MRQty > 0 Then
                
       
                
                    StockQTY = CheckDot(Trim(!StockQTY))
                    
                   If bTutti = False Then If StockQTY < MRQty Then GoTo cont
                
            End If
            
            
            
        
            Grid.AddItem "", False
            
            ChemicalMRCode = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
            
            Grid.Cell(Grid.Rows - 1, 1).Text = ChemicalMRCode
            
            Grid.Cell(Grid.Rows - 1, 2).Text = GetSupplier(ChemicalMRCode, MNP)
            Grid.Cell(Grid.Rows - 1, 3).Text = MNP
            
            
            Grid.Cell(Grid.Rows - 1, 4).Text = IIf(IsNull(Trim(!Bottle)), "", Trim(!Bottle))
            Grid.Cell(Grid.Rows - 1, 5).Text = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
            Grid.Cell(Grid.Rows - 1, 6).Text = IIf(IsNull(Trim(!Lot)), "", Trim(!Lot))
            Grid.Cell(Grid.Rows - 1, 7).Text = PurityPerc & " %"
            
            
            
            Grid.Cell(Grid.Rows - 1, 8).Text = (IIf(IsNull(Trim(!MRValue)), "", Trim(!MRValue)))
            Grid.Cell(Grid.Rows - 1, 9).Text = IIf(IsNull(Trim(!Unit)), "", Trim(!Unit))
            Grid.Cell(Grid.Rows - 1, 10).Text = IIf(IsNull(Trim(!Location)), "", Trim(!Location))
            Grid.Cell(Grid.Rows - 1, 11).Text = CheckDot(IIf(IsNull(Trim(!StockQTY)), 0, !StockQTY))
            Grid.Cell(Grid.Rows - 1, 11).BackColor = vbColorResults
            Grid.Cell(Grid.Rows - 1, 12).Text = IIf(IsNull(Trim(!stockUnit)), "mL", Trim(!stockUnit))
            Grid.Cell(Grid.Rows - 1, 13).Text = FormatDataLAT(IIf(IsNull(Trim(!ArrivedTime)), "", Trim(!ArrivedTime)))
            Grid.Cell(Grid.Rows - 1, 14).Text = FormatDataLAT(IIf(IsNull(Trim(!Open)), "", Trim(!Open)))
            Grid.Cell(Grid.Rows - 1, 15).Text = FormatDataLAT(IIf(IsNull(Trim(!Finished)), "", Trim(!Finished)))
            Grid.Cell(Grid.Rows - 1, 16).Text = FormatDataLAT(IIf(IsNull(Trim(!SupplierEXP)), "", Trim(!SupplierEXP)))
            Grid.Cell(Grid.Rows - 1, 17).Text = FormatDataLAT(IIf(IsNull(Trim(!MREXP)), "", Trim(!MREXP)))
            Grid.Cell(Grid.Rows - 1, 18).Text = GetStatus(IIf(IsNull(Trim(!Status)), 0, Trim(!Status)))
            Grid.Cell(Grid.Rows - 1, 19).Text = IIf(IsNull(Trim(!Note)), "", Trim(!Note))
            
            Grid.Cell(Grid.Rows - 1, 20).Text = !ID
            
             Grid.Cell(Grid.Rows - 1, 21).Text = IIf(IsNull(Trim(!U)), "", Trim(!U))
            
       
            
            
            Call GetDatabaseWareHouseEntry(!ID, i, False, iWarehouse)
       
          
         
            
            Grid.Cell(Grid.Rows - 1, 3).Alignment = cellLeftCenter
            
           
            Grid.Cell(Grid.Rows - 1, 5).Alignment = cellLeftCenter
            Grid.Cell(Grid.Rows - 1, 6).Alignment = cellLeftCenter
            
            Grid.Cell(Grid.Rows - 1, 8).Alignment = cellRightCenter
            Grid.Cell(Grid.Rows - 1, 9).Alignment = cellLeftCenter
            
            Grid.Cell(Grid.Rows - 1, 11).Alignment = cellRightCenter
            Grid.Cell(Grid.Rows - 1, 12).Alignment = cellLeftCenter
            
            Grid.Cell(Grid.Rows - 1, 18).Alignment = cellCenterCenter
            Grid.Cell(Grid.Rows - 1, 4).FontBold = True
            Grid.Cell(Grid.Rows - 1, 4).ForeColor = &H473733
            Grid.Cell(Grid.Rows - 1, 6).FontBold = True
            Grid.Cell(Grid.Rows - 1, 6).ForeColor = &H473733
            
            
            Grid.Cell(Grid.Rows - 1, 11).FontBold = True
            Grid.Cell(Grid.Rows - 1, 11).ForeColor = &H473733
            Grid.Cell(Grid.Rows - 1, 12).FontBold = True
            Grid.Cell(Grid.Rows - 1, 12).ForeColor = &H473733
            
            
          '  Grid.Cell(Grid.Rows - 1, 6).BackColor = vbColorResults
            
            Select Case Grid.Cell(Grid.Rows - 1, 18).Text
                Case "In Stock"
                    Grid.Cell(Grid.Rows - 1, 18).BackColor = vbColorGreen
                    Grid.Cell(Grid.Rows - 1, 18).ForeColor = vbWhite
                Case "Opened"
                    For t = 1 To Grid.Cols - 1
                        Grid.Cell(Grid.Rows - 1, t).ForeColor = vbColorBlueProgram
                        Grid.Cell(Grid.Rows - 1, t).FontBold = True
                    Next
                    
                    Grid.Cell(Grid.Rows - 1, 18).BackColor = vbColorBlueProgram
                    Grid.Cell(Grid.Rows - 1, 18).ForeColor = vbWhite
                Case Else
                    For t = 1 To Grid.Cols - 1
                        Grid.Cell(Grid.Rows - 1, t).ForeColor = vbColorRed
                        Grid.Cell(Grid.Rows - 1, t).FontBold = True
                    Next
                    
                    Grid.Cell(Grid.Rows - 1, 18).BackColor = vbColorRed
                    Grid.Cell(Grid.Rows - 1, 18).ForeColor = vbWhite
                
                 
            End Select
            
            
            
             If IsDate(Grid.Cell(Grid.Rows - 1, 16).Text) Then
                Dim SuppExpDate As Date
                SuppExpDate = Grid.Cell(Grid.Rows - 1, 16).Text
              
                
                If SuppExpDate < FormatDateTime(Now(), vbShortDate) Then
                
                    For x = 1 To Grid.Cols - 1
                       ' Grid.Cell(Grid.Rows - 1, x).BackColor = vbColorRed
                        Grid.Cell(Grid.Rows - 1, x).ForeColor = vbColorRed
                    Next
                    Grid.Cell(Grid.Rows - 1, 18).BackColor = vbColorRed
                    Grid.Cell(Grid.Rows - 1, 18).ForeColor = vbWhite
                  '  Grid.Cell(Grid.Rows - 1, 18).ForeColor = vbColorRosaTabella
                    
                ElseIf SuppExpDate = FormatDateTime(Now(), vbShortDate) Then
                
                    Grid.Cell(Grid.Rows - 1, 17).BackColor = vbColorOrange
                    Grid.Cell(Grid.Rows - 1, 17).ForeColor = vbWhite
                    Grid.Cell(Grid.Rows - 1, 18).BackColor = vbColorOrange
                    Grid.Cell(Grid.Rows - 1, 18).ForeColor = vbWhite
                    
                End If
               
             
             End If
            
            If IsDate(Grid.Cell(Grid.Rows - 1, 17).Text) Then
               
                Dim ExpDate As Date
                
                ExpDate = Grid.Cell(Grid.Rows - 1, 17).Text
              
                
            
                If ExpDate < FormatDateTime(Now(), vbShortDate) Then
                
                    For x = 1 To Grid.Cols - 1
                       ' Grid.Cell(Grid.Rows - 1, x).BackColor = vbColorRed
                        Grid.Cell(Grid.Rows - 1, x).ForeColor = vbColorRed
                    Next
                    Grid.Cell(Grid.Rows - 1, 18).BackColor = vbColorRed
                    Grid.Cell(Grid.Rows - 1, 18).ForeColor = vbWhite
                  '  Grid.Cell(Grid.Rows - 1, 18).ForeColor = vbColorRosaTabella
                    
                ElseIf ExpDate = FormatDateTime(Now(), vbShortDate) Then
                
                    Grid.Cell(Grid.Rows - 1, 17).BackColor = vbColorOrange
                    Grid.Cell(Grid.Rows - 1, 17).ForeColor = vbWhite
                    Grid.Cell(Grid.Rows - 1, 18).BackColor = vbColorOrange
                    Grid.Cell(Grid.Rows - 1, 18).ForeColor = vbWhite
                    
                End If
               
                
            End If
            
            ' calcolo la stock qty dell'RM
            
            If (IsNull(Trim(!StockQTY)) Or Trim(!StockQTY) = "") And Not (!bClosed) And ChemicalMRCode <> "" Then
            
            
            Else
                
                ' sommo la qty
                  StockQTY = (Trim(!StockQTY))
                  stockUnit = IIf(IsNull(Trim(!stockUnit)), "mL", Trim(!stockUnit))
                  
                  Call SetChemicalMRStockQTY(ChemicalMRCode, StockQTY, stockUnit, MRCode)
                  
            End If
            
            
            
cont:
            .MovePrevious
        Next
    
    End With
    For i = 1 To .Cols - 4
        .Column(i).AutoFit
        .Column(i).Width = .Column(i).Width * 1.1
    Next
    .Column(5).Width = 200
    .Column(20).Width = 0
    .AllowUserReorderColumn = True
    
    .Column(14).Sort cellDescending
    
    
    
   
    .Refresh
    .AutoRedraw = True
    .ReadOnly = True


End With


    ' aggiorno DB MR Stock Qty

    Call UpdateStockQTY_MRDatabase(MRCode)

End Function




Public Function MRSearchInGrid(ByRef Grid As Grid, ByVal str As String, ByVal bShowAll As Boolean)

Dim i As Integer

str = Trim(str)

If str = "" Then bShowAll = True

With Grid
    .AutoRedraw = False
    If .Rows > 1 Then
        
        For i = 1 To .Rows - 1
            If bShowAll Then
                 .RowHeight(i) = 25
            Else
                If InStr(UCase(.Cell(i, 1).Text), UCase(str)) Then
                    
                    .RowHeight(i) = 25
                
                Else
                    .RowHeight(i) = 0
                
                End If
            End If
            
        
        Next

    End If
    .Refresh
    .AutoRedraw = True
End With
End Function

Public Function GetSupplier(ByVal MRCode As String, ByRef MNP As String) As String

If MRCode <> "" Then

    With dbTabMR
        .filter = ""
        .filter = "Code='" & MRCode & "'"
        If .EOF Then
        Else
                  GetSupplier = IIf(IsNull(Trim(!Supplier)), "", Trim(!Supplier))
                  MNP = IIf(IsNull(Trim(!MNP)), "", Trim(!MNP))
        
        End If
    End With
End If

End Function


