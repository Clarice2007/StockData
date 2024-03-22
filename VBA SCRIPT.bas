Sub Stockdata()
     'setting up the Variables
        Dim tickername As String
        Dim tickervolume As Double
        Dim summary_ticker_row As Integer
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim summaryTickerRow As Integer
        Dim greatestIncreaseTicker As String
        Dim greatestDecreaseTicker As String
        Dim greatestVolumeTicker As String
        Dim greatestIncrease As Double
        Dim greatestDecr As Double
        Dim greatestVol As Double

        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        tickervolume = 0
        summary_ticker_row = 2

        
        'Headers for Summary Table
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
      
        'calculate Open Price
        open_price = Cells(2, 3).Value

        'Count of the rows .
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through the rows

        For i = 2 To lastrow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
              'the ticker name
              tickername = Cells(i, 1).Value

              ' volume of trade
              tickervolume = tickervolume + Cells(i, 7).Value

              'Print the ticker name
              Range("I" & summary_ticker_row).Value = tickername

              'Print the trade volume
              Range("L" & summary_ticker_row).Value = tickervolume

              close_price = Cells(i, 6).Value

              ' yearly change
              yearly_change = (close_price - open_price)
              
              'Print the yearly change
              Range("J" & summary_ticker_row).Value = yearly_change

             'when calculating the percent change
                If (open_price = 0) Then

                    percent_change = 0

                Else
                    
                    percent_change = yearly_change / open_price
                
                End If

              Range("K" & summary_ticker_row).Value = percent_change
              Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              summary_ticker_row = summary_ticker_row + 1

              tickervolume = 0

              open_price = Cells(i + 1, 3)
            
            Else
              
    
              tickervolume = tickervolume + Cells(i, 7).Value

            
            End If
        
        Next i

    lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color coding for the year change
        For i = 2 To lastrow_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i

        

    'the max and min values Calculations
        For i = 2 To lastrow_summary_table
            'Calculate the max
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
                Cells(2, 16).Value = Cells(i, 9).Value
                Cells(2, 17).Value = Cells(i, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"

            'Calculate the min
            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
                Cells(3, 16).Value = Cells(i, 9).Value
                Cells(3, 17).Value = Cells(i, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            
            'Find the maximum volume of trade
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
                Cells(4, 16).Value = Cells(i, 9).Value
                Cells(4, 17).Value = Cells(i, 12).Value
            
            End If
        
        Next i
        End Sub
