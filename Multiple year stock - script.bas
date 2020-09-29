Attribute VB_Name = "Module1"
Sub tickerloop()

      For Each WS In Worksheets
      
        'Declare the stock calculation variables

        Dim ticker As String
        Dim counter, total As Variant
        Dim year_open, year_close As Variant
        Dim last_row, last_ticker As Variant
        Dim yearly_change, percent_change As Variant
        
            WS.Range("I1").Value = "Ticker"
            WS.Range("J1").Value = "Yearly Change"
            WS.Range("K1").Value = "Percent Change"
            WS.Range("L1").Value = "Total Stock Volume"
        
            'Find the last row of column 1
        last_row = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
         'Assign variable values
        
        ticker = ""
        counter = 1
        total = 0
        year_open = 0
        year_close = 0
        
            'For all the rows that have a value in column 1
        For Row = 2 To last_row:
        
                'Count a new ticker name
            If WS.Cells(Row, 1).Value <> ticker Then
                       
                'Get a new name and set current Ticker to new Ticker name
                counter = counter + 1
                ticker = WS.Cells(Row, 1).Value
                
                'Assign a new row in the Ticker Column with the Ticker name
                WS.Cells(counter, 9).Value = ticker
                
                'Set the Year Open to the first entry's Opening number
                year_open = WS.Cells(Row, 3).Value
                
                'Set the starting Total to the first Volume record
                total = WS.Cells(Row, 7).Value
                WS.Cells(counter, 12) = total
            
              Else
                'If Ticker is the same, print the Volume to the Total
                total = total + WS.Cells(Row, 7).Value
                WS.Cells(counter, 12).Value = total
                    
            End If
            
                'If it's the last entry for Ticker, set Year Close value
            If WS.Cells((Row + 1), 1).Value <> ticker Then
            year_close = WS.Cells(Row, 6).Value
            
                'Calculate the Yearly Change from Open to Close
                yearly_change = year_close - year_open
                    
                'Print to Yearly Change Column
                WS.Cells(counter, 10).Value = yearly_change
                
                'Calculate Percent Change from the Yearly Open
                If year_open = 0 Then
                    percent_change = yearly_change
                Else
                    percent_change = yearly_change / year_open
                End If
                    
                'Print to Percent Change Column
                WS.Cells(counter, 11).Value = percent_change
                
                'Color cells based on positive or negative change
                If yearly_change < 0 Then
                    WS.Cells(counter, 10).Interior.ColorIndex = 3
                
                ElseIf yearly_change > 0 Then
                    WS.Cells(counter, 10).Interior.ColorIndex = 4
                
                End If
            
            End If
            
        Next Row
        
        'Declare the performance evaluation variables
        
        Dim top_increase As Variant
        Dim top_decrease As Variant
        Dim top_volume As Variant
        
        'Add columns labels
        
        WS.Range("O2").Value = "Greatest % Increase"
        WS.Range("O3").Value = "Greatest % Decrease"
        WS.Range("O4").Value = "Greatest Total Volume"
        WS.Range("P1").Value = "Ticker"
        WS.Range("Q1").Value = "Value"
        
        'Find the last row of list of Tickers
            last_ticker = WS.Cells(Rows.Count, 9).End(xlUp).Row
            
        'Assign variable values
        
        top_increase = 0
        top_decrease = 0
        top_total = 0
        
        For Row = 2 To last_ticker:
        
                'Set Positive Change leader
                If WS.Cells(Row, 11).Value > top_increase Then
                top_increase = WS.Cells(Row, 11).Value
                
                'Print Ticker and Percent into Summary
                WS.Range("P2").Value = WS.Cells(Row, 9).Value
                WS.Range("Q2").Value = WS.Cells(Row, 11).Value
            
            End If
            
                'Set Negative Change leader
                If WS.Cells(Row, 11).Value < top_decrease Then
                top_decrease = WS.Cells(Row, 11).Value
            
                'Print Ticker and Percent into Summary
                WS.Range("P3").Value = WS.Cells(Row, 9).Value
                WS.Range("Q3").Value = WS.Cells(Row, 11).Value
            
            End If
            
                'Set Total Stock Volume leader
                If WS.Cells(Row, 12).Value > top_total Then
                top_total = WS.Cells(Row, 12).Value
                
                'Print Ticker and Percent into Leaderboard
                WS.Range("P4").Value = WS.Cells(Row, 9).Value
                WS.Range("Q4").Value = WS.Cells(Row, 12).Value
            
            End If
        
        Next Row
        
    Next WS
    
End Sub
