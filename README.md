## Using VBA to Produce Stock Market Data

This script loops through each worksheet (i.e., year) and does the following:
- Creates table that displays ticker data aggregated over a year. Specifically, for each ticker table contains:
      - Ticker: Symbol of the ticker.
      - Yearly Change: Difference between the closing price at the end of the year and the opening price at the beginning of the year.
      - Percent Change: The percentage change from opening to closing price of the year; conditional formatting used to highlight positive changes in green and negative in red.
      - Total Stock Volume: The total volume of the stock.
- Creates second table that returns the stock with the "Greatest % increase", "Greatest % decrease," and "Greatest total volume".

      Code Snippet:
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
