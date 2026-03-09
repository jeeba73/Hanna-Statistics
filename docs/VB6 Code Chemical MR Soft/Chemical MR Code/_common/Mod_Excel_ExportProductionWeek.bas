Attribute VB_Name = "Mod_Excel_ExportProductionWeek"
Option Explicit

Private uSTDPreparation As RecipeForSTDPreparation
Private uProdWeekFileArray() As String
Private uProdCount As Integer
Private uWeekProd As String
Private uProdLine As String

Public Function EsportaSTDPreparationWeekExcel(ByRef ProdWeekFileArray() As String, ByVal sString As String, ByVal ProdCount As Integer, ByVal WeekProd As String, ByVal ProdLine As String) As Boolean
    Dim rc As Boolean
    On Error GoTo ERR_IMP
    rc = True
    
    uProdCount = ProdCount
    uProdWeekFileArray = ProdWeekFileArray
    uWeekProd = WeekProd
    uProdLine = ProdLine

        If CreateExcel(False) Then
            NewExcelWorksheet (sString)
            If CopySTDPreparationData Then
                Call SaveExcel(sString)
                Call CloseExcel
                PopupMessage 2, "Excel file correctly generated..."
            Else
                rc = False
            End If
        Else
            rc = False
        End If
ERR_END:
    On Error GoTo 0
    EsportaSTDPreparationWeekExcel = rc
    Exit Function
ERR_IMP:
    rc = False
    MsgBox err.Description
    Resume ERR_END
End Function


Private Function CopySTDPreparationData() As Boolean
Dim rc As Boolean
Dim i As Integer
    On Error GoTo ERR_COPY
    '---------------------------
    ' set excel page
    '---------------------------
   ' Call SetUnit
    Call FormatPage
    Call SetInformation(i)
    Call SetHannaCode(i)

    
    rc = True
ERR_END:
    On Error GoTo 0
    CopySTDPreparationData = rc
    Exit Function
ERR_COPY:
    rc = False
    MsgBox err.Description
    Resume Next
End Function

Private Sub SetInformation(ByRef Riga As Integer)
Dim i As Integer
Dim sString As String
Dim rc As Boolean
Riga = 2


    Call AddValue(Riga + 2, 2, "STDPreparation per Week", True, True)

    
  
    Call AddValue(Riga + 5, 2, "Line", True)
    Call AddValue(Riga + 5, 3, "Week STDPreparation", True)
    Call AddValue(Riga + 6, 2, uProdLine, True)
    Call AddValue(Riga + 6, 3, "'" & uWeekProd, True)
   
  


    Riga = Riga + 10
 
End Sub
Private Sub SetHannaCode(ByRef Riga As Integer)
Dim i As Integer
Dim t As Integer
Dim HannaCount As Integer
Dim Variance As String
Dim VarDbl As Double
Dim PercStr As String
Dim ProdTotal As Double

    Riga = Riga + 2
    
    Call AddValue(Riga - 1, 2, "Hanna Code Table", True, True)




    'uProdCount = ProdCount
    'uProdWeekFileArray = ProdWeekFileArray
    'uWeekProd = WeekProd
    
    
    ProdTotal = 0

On Error GoTo ERR_GET:
        '------------------------------------------------
        '      STDPreparation  TABELLA HANNA CODE
        '------------------------------------------------
        '.Cell(0, 1).Text = "Code"
        '.Cell(0, 2).Text = "Product Name"
        '.Cell(0, 3).Text = "Line"
        '.Cell(0, 4).Text = "Volume/Weight"
        '.Cell(0, 5).Text = "(um)"
        '.Cell(0, 6).Text = "Q.ty to produce"
        '.Cell(0, 7).Text = "Q.ty to produced"
        '.Cell(0, 8).Text = ""
        
        '.Cell(0, 9).Text = "%"
        '.Cell(0, 10).Text = "Recipe"
        '.Cell(0, 11).Text = "Mix"
        
        Call AddValue(Riga, 2, "Code", True)
        Call AddValue(Riga, 3, "Product Name", True)
        Call AddValue(Riga, 4, "STDPreparation Date", True)
        Call AddValue(Riga, 5, "Volume/Weight", True)
        Call AddValue(Riga, 6, "(um)", True)
        Call AddValue(Riga, 7, "Q.ty to produce", True)
        Call AddValue(Riga, 8, "Q.ty  produced", True)
        Call AddValue(Riga, 9, "%", True)
        Call AddValue(Riga, 10, "Recipe", True)
        
        If InStr(LCase(uProdLine), "all lines") Then
        
            Call AddValue(Riga, 11, "Line", True)
        End If
       
    For t = 1 To uProdCount
    
        ' acquisisco uSTDPreparation
        If uProdWeekFileArray(t) = "" Then GoTo contProd
        Call STDPreparationGetSetting(uSTDPreparation, uProdWeekFileArray(t))
        With uSTDPreparation
        
            
            HannaCount = .HannaCodesCount
            
            For i = 1 To HannaCount
            
                If .HannaCodes(i).bHide Then GoTo cont
                If (.HannaCodes(i).QtyToProduce) = "" And (.HannaCodes(i).QtyProduced) = "" Then GoTo cont
                If CDbl(.HannaCodes(i).QtyToProduce) = 0 And CDbl(.HannaCodes(i).QtyProduced) = 0 Then GoTo cont
                
                
                Riga = Riga + 1
                
                
                Call AddValue(Riga, 2, .HannaCodes(i).Code)
                Call AddValue(Riga, 3, .HannaCodes(i).ProductName)
                Call AddValue(Riga, 4, .HannaCodes(i).DateProd)
                Call AddValue(Riga, 5, Replace(.HannaCodes(i).Qty, ",", "."))
                Call AddValue(Riga, 6, .HannaCodes(i).Um)
                Call AddValue(Riga, 7, Replace(.HannaCodes(i).QtyToProduce, ",", "."))
                Call AddValue(Riga, 8, Replace(.HannaCodes(i).QtyProduced, ",", "."))
                

                If .HannaCodes(i).QtyToProduce = "" Then .HannaCodes(i).QtyToProduce = "0"
                If .HannaCodes(i).QtyProduced = "" Then .HannaCodes(i).QtyProduced = "0"
                
                
                ProdTotal = ProdTotal + CDbl(.HannaCodes(i).QtyProduced)
                If CDbl(.HannaCodes(i).QtyProduced) > 0 And CDbl(.HannaCodes(i).QtyToProduce) > 0 Then
                
                    VarDbl = FormatNumber((.HannaCodes(i).QtyProduced / .HannaCodes(i).QtyToProduce), 4) * 100
                     
                    Select Case VarDbl
                        Case Is < 100
                            PercStr = "'- "
                            VarDbl = FormatNumber(100 - VarDbl, 2)
                        Case Is = 100
                            PercStr = ""
                            VarDbl = VarDbl
                        Case Is > 100
                            PercStr = "'+ "
                            VarDbl = VarDbl
                    End Select
                                       
                    Variance = PercStr & VarDbl & " %"

                    Call AddValue(Riga, 9, Replace(Variance, ",", "."))
                    
                    VarDbl = CDbl(.HannaCodes(i).QtyProduced) - CDbl(.HannaCodes(i).QtyToProduce)
                Else
                    Call AddValue(Riga, 9, "/")
                    
                End If
                
                If CDbl(.HannaCodes(i).QtyToProduce) = 0 Then
                    VarDbl = CDbl(.HannaCodes(i).QtyProduced)
                End If
                Call AddValue(Riga, 10, .HannaCodes(i).Recipe)
                If InStr(LCase(uProdLine), "all lines") Then
                    Call AddValue(Riga, 11, .HannaCodes(i).Line)
                End If
        
cont:
            Next
        
        
        End With
        
contProd:

    Next
ERR_END:

    Riga = Riga + 4
    Call AddValue(Riga - 1, 8, "Total Q.ty", True)
    Call AddValue(Riga, 8, "'" & ProdTotal)
    
    Exit Sub
ERR_GET:
   MsgBox err.Description
   Resume Next

End Sub


Private Sub SetAcquisitionGrid(ByRef Riga As Integer)
Dim i As Integer
Dim t As Integer
Dim RecipeCount As Integer
Dim Variance As Double
Dim VariancePerc As Double
Dim TotalRealWeight As Double
Dim bUmMassa As Boolean
Dim Density As Double
Dim bRecalculate As Boolean
Dim PesoIntolleranza As Double
Dim MyColor As OLE_COLOR

Dim iAcquisition As PrepAcquisition
 
On Error GoTo ERR_GET:


    Riga = Riga + 2
    
    Call AddValue(Riga - 1, 2, "Acquisition Table", True, True)
    
    '------------------------------------------------
    '      Acquisition Grid
    '------------------------------------------------
    
    Call AddValue(Riga, 2, "Code", True)
    Call AddValue(Riga, 3, "QtyProduced", True)
    Call AddValue(Riga, 4, "LotNumber", True)
    Call AddValue(Riga, 5, "Operator", True)
    Call AddValue(Riga, 6, "DateProd", True)
    Call AddValue(Riga, 7, "WeekProd", True)
    Call AddValue(Riga, 8, "Machine", True)
    Call AddValue(Riga, 9, "Note", True)
    Call AddValue(Riga, 10, "AcquisitionTime", True)
    Call AddValue(Riga, 11, "Mix1Lot", True)
    Call AddValue(Riga, 12, "Mix2Lot", True)
    Call AddValue(Riga, 13, "Exp Date", True)
            
     For i = 1 To UBound(uSTDPreparation.HannaCodes)

        With uSTDPreparation.HannaCodes(i)
          
          
          
            If .AcquisitionCount > 0 Then
                For t = 1 To .AcquisitionCount
                    Call STDPreparationAddNewRowInAcquisitionExcel(Riga, .Acquisitions(t))
                Next
            End If
        End With
    Next
    



ERR_END:
   Riga = Riga + 2
   On Error GoTo 0
    Exit Sub
ERR_GET:
    MsgBox err.Description
    Resume Next
End Sub

Private Function STDPreparationAddNewRowInAcquisitionExcel(ByRef Riga As Integer, ByRef iAcquisition As ProdAcquisition)
                
                
               
                
                    Riga = Riga + 1
                    
                   
                    
                    
                    Call AddValue(Riga, 2, iAcquisition.Code)
                    Call AddValue(Riga, 3, CStr(Replace(iAcquisition.QtyProduced, ",", ".")))
                    Call AddValue(Riga, 4, iAcquisition.LotNumber)
                    Call AddValue(Riga, 5, iAcquisition.Operator)
                    Call AddValue(Riga, 6, FormatDataLAT(iAcquisition.DateProd))
                    Call AddValue(Riga, 7, iAcquisition.WeekProd)
                    Call AddValue(Riga, 8, iAcquisition.Machine)
                    Call AddValue(Riga, 9, iAcquisition.Note)
                    Call AddValue(Riga, 10, CStr(iAcquisition.AcquisitionTime))
                    Call AddValue(Riga, 11, iAcquisition.Mix1Lot)
                    Call AddValue(Riga, 12, iAcquisition.Mix2Lot)
                    Call AddValue(Riga, 13, "'" & iAcquisition.ExpDate)
                    
                    
                    
                    
End Function

