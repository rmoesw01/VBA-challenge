Attribute VB_Name = "Module2"
Sub VBAStocks2():
    ' A script that will loop through all the stocks for one year and output the following information.
    '       The ticker symbol.
    '       Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    '       The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    '       The total stock volume of the stock.

    ' Data File Format requirements
    ' column A: ticker symbol, column B: date, column C: opening value, column F: closing value, column G: stock volume
    ' columns H - L should be blank
    ' Data must be sorted first by ticker column, then by date column (smallest to largest)
    ' **** CAUTION **** This program does not allow for skipped rows, please make sure your data is all together
    
    ' Initialize variables for finding greatest percent increase, percent decrease, & total stock volume
    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0
    
    'Print output headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Initialize output row & first opening value
    output_row = 2
    open_value = Cells(2, 3).Value
    
    'Loop through all of the rows of the spreadsheet that contain data (no skipped rows allowed)
    For i = 2 To Range("A1").End(xlDown).Row + 1
        
        ' Check to see if the next row in the data has a different ticker symbol (data mush be sorted by ticker column first, then by date)
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            ' Grab the closing value for the year
            close_value = Cells(i, 6).Value
            
            ' Calculate the change for the year and the percent change for the year
            year_change = close_value - open_value
            If open_value = 0 Then
                percent_change = 0
            Else
                percent_change = year_change / open_value
            End If
            
            ' Add the final volume value to the total stock volume
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
            
            ' Print output on current output row
            Cells(output_row, 9).Value = Cells(i, 1).Value
            Cells(output_row, 10).Value = year_change
            Cells(output_row, 11).Value = percent_change
            Cells(output_row, 12).Value = total_stock_volume
            
            ' Check to see if the values are the greatest percent increase, percent decrease, & total stock volume
            If greatest_increase < percent_change Then
                greatest_increase = percent_change
                greatest_increase_ticker = Cells(i, 1).Value
            ElseIf greatest_decrease > percent_change Then
                greatest_decrease = percent_change
                greatest_decrease_ticker = Cells(i, 1).Value
            End If
            
            If greatest_volume < total_stock_volume Then
                greatest_volume = total_stock_volume
                greatest_volume_ticker = Cells(i, 1).Value
            End If
            
            ' Conditional formatting that will highlight positive change in green and negative change in red.
            If year_change > 0 Then
                Cells(output_row, 10).Interior.ColorIndex = 4
            ElseIf year_change < 0 Then
                Cells(output_row, 10).Interior.ColorIndex = 3
            End If
                        
            ' Reset variable values for next ticker symbol
            open_value = Cells(i + 1, 3).Value
            total_stock_volume = Cells(i + 1, 7).Value
            output_row = output_row + 1
            
        Else
            If open_value = 0 Then
                open_value = Cells(i, 3).Value
            End If
            
            ' Add the current row's stock volume to the total stock volume
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
            
        End If
    Next i

    ' output greatest percent increase, percent decrease, & total stock volume values
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    Range("P1").Value = "Ticker"
    Range("P2").Value = greatest_increase_ticker
    Range("P3").Value = greatest_decrease_ticker
    Range("P4").Value = greatest_volume_ticker
    
    Range("Q1").Value = "Value"
    Range("Q2").Value = greatest_increase
    Range("Q3").Value = greatest_decrease
    Range("Q4").Value = greatest_volume

    ' Format Output Display
    Range("I:Q").Columns.AutoFit
    Range("I:Q").HorizontalAlignment = xlHAlignCenter
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = "#,###"
    Range("K:K").NumberFormat = "0.00%"
    Range("L:L").NumberFormat = "#,###"

End Sub

