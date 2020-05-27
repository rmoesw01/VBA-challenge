Attribute VB_Name = "Module1"
Sub VBAStocks():
' A script that will loop through all the stocks for one year and output the following information.
'       The ticker symbol.
'       Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'       The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'       The total stock volume of the stock.

' **** CAUTION **** This program does not allow for skipped rows, please make sure your data is all together

'-------------   Search through first column to look for unique values   -------------
    'Declare array to hold unique ticker names
    Dim Tickers(1000) As String
    
    'Declare & Initialize variable to count how many unique ticker names there are
    Dim TickerCount As Integer
    TickerCount = 0
    
    'Declare & Initialize variable to count how many rows have data in the first column
    Dim rowcount As Long
    rowcount = 0
    
    'Declare loop counters
    Dim rows As Long
    Dim i As Integer
    
    'Declare variable to keep track of if a cell value is already in the tickers array
    Dim found As Boolean
    
    'Loop through all of the rows of the spreadsheet
    For rows = 1 To Range("A:A").Cells.Count
        
        'reset found variable for new cell value check
        found = 0
        
        'Check to see if the current row has a value in the first column
        If Not (IsEmpty(Cells(rows, 1))) Then
            'Increase the count of how many rows have values in the first column
            rowcount = rowcount + 1
            
            'Check to see if current value is already in the Tickers array
            For i = 0 To TickerCount
                If Cells(rows, 1).Value = Tickers(i) Then
                    'indicate that the value was found and stop searching
                    found = 1
                    Exit For
                End If
            Next i
            
            'Add the value to the array because it was not found
            If Not (found) Then
                Tickers(TickerCount) = Cells(rows, 1).Value
                TickerCount = TickerCount + 1
            End If
        End If
    Next rows
    
    'Print the header and and all the values in the tickers array
    Range("I1").Value = "Ticker"
    For i = 1 To TickerCount
        Cells(i + 1, 9).Value = Tickers(i)
    Next i
'-------------   End Search through first column to look for unique values   -------------

'   Look for the Min & max values in the date column and find the difference of the opening price and closing price for those two values for each ticker value
    'Declare variables to hold the minimum and maximum date values for the current ticker
    Dim startDate As Long
    Dim endDate As Long
    
    'Declare variables to hold the row reference of the start and end dates for the current ticker
    Dim startrow As Long
    Dim endrow As Long
    Dim stockvolume As Double
    
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    startrow = 2
    'Loop through each ticker value
    For i = 1 To TickerCount

        startDate = Cells(startrow, 2).Value
        endDate = Cells(startrow, 2).Value
        stockvolume = 0
        
        For rows = startrow + 1 To rowcount
            If Cells(rows, 1).Value = Tickers(i) Then
                If Cells(rows, 2).Value > endDate Then
                    endDate = Cells(rows, 2).Value
                    endrow = rows
                    
                    '   Add all the values in the vol column for the ticker
                    stockvolume = stockvolume + Cells(rows, 7).Value
                End If
            Else
                Debug.Print "i = " & i & ", startrow = " & startrow & ", endrow = " & endrow
                Cells(i + 1, 10).Value = Cells(endrow, 6).Value - Cells(startrow, 3).Value
                
                'Calculate percent change by dividing the yearly change by the opening price at the beginning of the year
                Cells(i + 1, 11).Value = (Cells(endrow, 6).Value - Cells(startrow, 3).Value) / Cells(startrow, 3).Value
                
                'Display total stock volume for the current ticker
                Cells(i + 1, 12).Value = stockvolume
                
                startrow = endrow + 1
                Exit For
            End If
        Next rows
        
        'MsgBox ("startrow = " & startrow)
    Next i

'   Format Output Display
    Range("K:K").NumberFormat = "0.00%"
    Range("L:L").NumberFormat = "#,###"

' Conditional formatting that will highlight positive change in green and negative change in red.

End Sub
