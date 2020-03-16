Sub YearlyStocks()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

        ' Set an initial variable for holding the stock tickers
        Dim Stocks_Ticker As String

        ' Set an initial variable for opening price at the beginning of a year
        Dim Stocks_OpenPrice As Double
        
        ' Set an initial variable for Closing price at the end of a year
        Dim Stocks_ClosePrice As Double
        
        ' Set an initial variable for Volume Change
        Dim Stocks_VolChg As Double

        ' Set an initial variable for holding the stock pct change
        Dim Stocks_PctChg As Double

         ' Set an initial variable for holding the stock total volume
        Dim Stocks_TotalVol As Double
        Stocks_TotalVol = 0

        ' Determine the Last Column Number
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' MsgBox (LastRow)
      
        ' Keep track of the location for each stock ticker in the summary table
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        ' Keep track of the location for each stock ticker in the summary table
        Dim Summary_Table_Row1 As Long
        Summary_Table_Row1 = 2
        
        
            ' Loop through all the stocks
            For i = 2 To LastRow
            
               ' Flagging when a Stock Ticker Changes
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                    ' Set the New Stock Ticker name
                    Stocks_Ticker = ws.Cells(i, 1).Value
                
                    ' Print the Stock Ticker to the Summary Table
                    ws.Range("I" & Summary_Table_Row).Value = Stocks_Ticker
                
                    Stocks_ClosePrice = ws.Cells(i, 6).Value
                    ws.Range("K" & Summary_Table_Row).Value = Stocks_ClosePrice
                    
                    ' Add to the Volume Total
                    Stocks_TotalVol = Stocks_TotalVol + ws.Cells(i, 7).Value
                    
                    ' Print the Volume Amount to the Summary Table
                    ws.Range("N" & Summary_Table_Row).Value = Stocks_TotalVol
                    
                          ' Reset the Brand Total
                    Stocks_TotalVol = 0
                
                    ' Add one to the summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
                    
               ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                
                    Stocks_OpenPrice = ws.Cells(i, 3).Value
                    ws.Range("J" & Summary_Table_Row1).Value = Stocks_OpenPrice
                    
                      ' Add one to the summary table row
                    Summary_Table_Row1 = Summary_Table_Row1 + 1
                    
                    ' Add to the Stock Total
                Stocks_TotalVol = Stocks_TotalVol + ws.Cells(i, 7).Value
                    
                 Else
                 
               ' Add to the Stock Total
                Stocks_TotalVol = Stocks_TotalVol + ws.Cells(i, 7).Value
                    
            
                End If
            
            Next
            
            ' Determine the Last Column for Ticker
        LastRowTicker = ws.Cells(Rows.Count, 9).End(xlUp).Row
        ' MsgBox (LastRowTicker)
        
           For t = 2 To LastRowTicker
           
              ' Obtain the Volume Change
              Stocks_VolChg = (ws.Cells(t, 11).Value - ws.Cells(t, 10).Value)
              ws.Cells(t, 12).Value = Format(Stocks_VolChg, "Fixed")
              

            
               If ws.Cells(t, 10).Value <> 0 Then
            
             ' Obtain the Volume % Change
              Stocks_PctChg = ((ws.Cells(t, 11).Value / ws.Cells(t, 10).Value) - 1)
              ws.Cells(t, 13).Value = Format(Stocks_PctChg, "Percent")
            
              End If
              
            ' Column Color for Volumn change
            If Stocks_PctChg > 0 Then
              
                ws.Cells(t, 12).Interior.ColorIndex = 4
                
            Else
                
                ws.Cells(t, 12).Interior.ColorIndex = 3
                
            End If
                                     
            
            Next
        
    'Deleting the Start and End stock values used in calculations above
    'PLEASE NOTE: RUNNING THIS SCRIPT TWICE WILL DELETE THE WRONG COLUMNS
    ws.Columns(10).EntireColumn.Delete
    ws.Columns(10).EntireColumn.Delete
    
    ' Print the Stock  Headers
     ws.Range("I1").Value = "Ticker"
     ws.Range("J1").Value = "Yearly Change"
     ws.Range("K1").Value = "Percent Change"
     ws.Range("L1").Value = "Total Stock Volume"

    Next
    

End Sub