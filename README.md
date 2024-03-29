## Using VBA to Analyze Stock Market Data

This script loops through each Excel worksheet (i.e., year) and does the following:
- Creates table that displays ticker data aggregated over a year. Specifically, for each ticker table contains:
     - Ticker: Symbol of the ticker.
     - Yearly Change: Difference between the closing price at the end of the year and the opening price at the beginning of the year.
     - Percent Change: The percentage change from opening to closing price of the year; conditional formatting used to highlight positive changes in green and negative changes in red.
     - Total Stock Volume: The total volume of the stock.
- Creates second table that returns the stock with the "Greatest % increase", "Greatest % decrease," and "Greatest total volume".

**Note:** Excel data not included.

### Code Snippet:
         For i = 2 To Rows.Count 'Looping thru rows
              If Cells(i, 11).Value <> "N/A" Then
                  If Cells(i, 11).Value > GreatestP Then
                      GreatestP = Cells(i, 11).Value
                      GreatestP_Ticker = Cells(i, 9).Value
                  ElseIf Cells(i, 11).Value < LeastP Then
                      LeastP = Cells(i, 11).Value
                      LeastP_Ticker = Cells(i, 9).Value
                  End If
              End If
          Next i

### Example of Worksheet After Running Code File
![Stock Prices](images/Stock(2016).PNG)
