'David Fried
'Data Science Bootcamp
'6/15/2020
'vba-challenge
'This script loops through each worksheet (i.e., year) and does the following:
'   Creates table that displays ticker data aggregated over a year. Specifically, for each ticker table contains-
'       Ticker: Symbol of the ticker.
'       Yearly Change: Difference between the closing price at the end of the year and the opening price at the beginning of the year.
'       Percent Change: The percentage change from opening to closing price of the year; conditional formatting used to highlight positive changes in green and negative in red.
'       Total Stock Volume: The total volume of the stock.
'   Creates second table that returns the stock with the "Greatest % increase", "Greatest % decrease," and "Greatest total volume".

Sub vba_challenge()
    Dim wksht As Integer
    Dim Ticker_Opening_Row As Long
    Dim Ticker_Closing_Row As Long
    Dim NT_Row As Integer 'Current Row for Table 1
    Dim Yearly_Change As Variant
    Dim Percent_Change As Variant
    Dim Total_Stock_Volume As Double
    Dim GreatestP as Double 'Greatest % increase
    Dim GreatestP_Ticker as String 'Ticker Name
    Dim GreatestTV as Double 'Greatest total volume
    Dim GreatestTV_Ticker as String 'Ticker Name
    Dim LeastP as Double 'Greatest % decrease
    Dim LeastP_Ticker as String 'Ticker Name
    
    wksht = 1
    For Each ws In Worksheets 'Looping thru worksheets
        Worksheets(wksht).Activate
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        'Ticker Opening Day Row (January 1)
        Ticker_Opening_Row = 2
        'Current row for Table 1
        NT_Row = 2
        
        'Creating Table 1 [Ticker_Year_Change, Percent Change, Total Stock Volume]
        For i = 2 To Rows.Count 'Looping thru rows
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            'Creates table values each time code identifies a transition between ticker values
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker_Closing_Row = i '(December 31)
                'Ticker Name
                Cells(NT_Row, 9).Value = Cells(i, 1).Value
                'Yearly Change & Percentage Change
                If Cells(Ticker_Opening_Row,3).Value = 0 Then
                    Yearly_Change = "N/A"
                    Percent_Change = "N/A"
                Else 
                    Yearly_Change = Cells(Ticker_Closing_Row, 6).Value - Cells(Ticker_Opening_Row, 3).Value
                    Percent_Change = Yearly_Change / Cells(Ticker_Opening_Row, 3)
                    Percent_Change = FormatPercent(Percent_Change, 2)
                End If
                Cells(NT_Row, 10).Value = Yearly_Change
                Cells(NT_Row, 11).Value = Percent_Change
                'Total Stock Volume
                Cells(NT_Row, 12).Value = Total_Stock_Volume
                'Changing variable values for next Table 1 entry
                Total_Stock_Volume = 0  
                Ticker_Opening_Row = i + 1 
                NT_Row = NT_Row + 1
            End If
            
            'End Table 1 if condition met
            If Cells(i + 1, 1).Value = "" Then
                Exit For 'End loop when reach blank cell in i,1
            End If 
            
        Next i
        
        'Table 1 conditional formatting For Percentage Change column
        'Green is positive; red is negative
        Range("K2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Font
            .Color = -16752384
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13561798
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Font
            .Color = -16383844
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13551615
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False

    'Creating Table 2 [Ticker, Value] from info in Table 1
    GreatestP = 0 'Greatest % increase
    GreatestP_Ticker = ""
    LeastP = 0 'greatest % decrease
    LeastP_Ticker = ""
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
    GreatestTV = 0
    GreatestTV_Ticker = ""
    For i = 2 To Rows.Count 'Looping thru rows
        If Cells(i, 12).Value > GreatestTV Then
            GreatestTV = Cells(i, 12).Value
            GreatestTV_Ticker = Cells(i, 9).Value
        End If
    Next i
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(2, 16).Value = GreatestP_Ticker
    Cells(2, 17).Value = FormatPercent(GreatestP, 2)
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(3, 16).Value = LeastP_Ticker
    Cells(3, 17).Value = FormatPercent(LeastP, 2)
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(4, 16).Value = GreatestTV_Ticker
    Cells(4, 17).Value = GreatestTV
    
    wksht = wksht + 1
    
    Next ws

End Sub

