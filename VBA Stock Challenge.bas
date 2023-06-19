Attribute VB_Name = "Module1"
Sub Stocks()

 For Each ws In Worksheets
 
    WorksheetName = ws.Name
    
    Dim Closing As Double
    Dim Opening As Double
    Dim Ticker As String
        
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Opening = ws.Cells(2, 3).Value
    Closing = 6
        
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Total_Volume As LongLong
    Dim Greatest_Total_Volume As LongLong
    Dim Summary_Table_Row As Integer
        
    Summary_Table_Row = 2
        
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percentage Change"
    ws.Cells(1, 13).Value = "Volume"
    ws.Columns("L").NumberFormat = "0.00%"
  
        
        
        
        
    Dim Summary_Table_Row2 As Integer
        
    Summary_Table_Row2 = 2
       
       'add summary table titles
        
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greates % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
  

 'loop through all rows
 
    For i = 2 To LastRow


            'check when ticker is different
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                            
                'grab the ticker, put in summary table
                                
                Ticker = ws.Cells(i, 1).Value
                
                'grab the ticker, put in summary table
                
                ws.Range("J" & Summary_Table_Row).Value = Ticker

                
                Yearly_Change = ws.Cells(i, Closing).Value - Opening
                
                
                ' store yeary change in summary table
                                
                
                ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
                
         If ws.Range("K" & Summary_Table_Row).Value > 0 Then
                    
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                    
         ElseIf ws.Range("K" & Summary_Table_Row).Value < 0 Then
                    
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                End If

                
                'Caluclate prenetage change
                
                
                Percent_Change = Yearly_Change / Opening
                
                
                ' Store Precentage Change Calculation
                
                ws.Range("L" & Summary_Table_Row).Value = Percent_Change

                'sum total volume per ticker
                
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
                'take total sum of volume and put in summary table
                
                ws.Range("M" & Summary_Table_Row).Value = Total_Volume
        
                'rest
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                Yearly_Change = 0
                                                
                Total_Volume = 0
                
                Opening = ws.Cells(i + 1, 3).Value
                
            
            Else
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
          End If
            
        Next i
            
         Greatest_Increase = Application.WorksheetFunction.Max(ws.Range("L:L"))
         Greatest_Decrease = Application.WorksheetFunction.Min(ws.Range("L:L"))
         Greatest_Total_Volume = Application.WorksheetFunction.Max(ws.Range("M:M"))
         ws.Range("R2", "R3").NumberFormat = "0.00%"
                 
    For i = 2 To LastRow
        If ws.Cells(i, 12).Value = Greatest_Increase Then
           ws.Range("R2").Value = Greatest_Increase
           ws.Range("Q2").Value = ws.Cells(i, 10).Value
        ElseIf ws.Cells(i, 12).Value = Greatest_Decrease Then
           ws.Range("R3").Value = Greatest_Decrease
           ws.Range("Q3").Value = ws.Cells(i, 10).Value
    End If
        If ws.Cells(i, 13).Value = Greatest_Total_Volume Then
           ws.Range("R4").Value = Greatest_Total_Volume
           ws.Range("Q4").Value = ws.Cells(i, 10).Value
    End If
            
        Next i
        
        
        ws.Columns("R").AutoFit
        ws.Columns("L").AutoFit
        ws.Columns("K").AutoFit
        ws.Columns("L").AutoFit
        ws.Columns("M").AutoFit
        ws.Columns("P").AutoFit
        ws.Columns("Q").AutoFit
Next ws

        
                
                
End Sub


