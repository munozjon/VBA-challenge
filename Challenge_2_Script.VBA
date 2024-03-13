Sub stocks_summary():
    dim i as long 'to be used to go through each row in the table
    dim current_ticker as string 'to be used for looping within one ticker at a time and output the ticker to the summary table
    dim year_open as double, year_close as double 'to be used for getting the open and close of the beginning/end of the year
    dim yearly_change as double, percent_change as double 'to be used for the summary table's formulas
    dim total_volume as double 'to be used to add each row's volume for each ticker and getting the total for each ticker
    dim j as long 'to be used to output totals for each ticker into their own rows in the summary table, iterating by 1 for each ticker change
    dim ws as Worksheet 'to use for iterating through each worksheet in the file

    'loop through each worksheet
    For Each ws In Worksheets

        'assign values to variables
        i = 2
        j = 2
        total_volume = 0
        year_open = 0
        year_close = 0

        'set column names for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'loop through the table while there is data in the row
        While Not IsEmpty(ws.Cells(i, 1).Value)
            
            'set current_ticker to the current_row's ticker and add the ticker to the summary table
            current_ticker = ws.Cells(i, 1).Value
            ws.Cells(j, 9).Value = current_ticker

            'loop through a ticker
            While ws.Cells(i, 1).Value = current_ticker

                'setting year_open to the first row's "open" of a ticker by using if statement to check if it's 0
                    'and if the "open" column has a value for that row
                    'this will prevent the variable from being overwrriten within the loop's iterations
                If year_open = 0 And Not IsEmpty(ws.Cells(i, 3).Value) Then
                    year_open = ws.Cells(i, 3).Value
                End if

                'add row's volume to the "total_volume" variable
                total_volume = total_volume + ws.Cells(i, 7).Value

                'move on to the next row by adding 1 to i
                i = i + 1

            wend

            '(i - 1) used to get the last row of the previous ticker, AKA the end of the year
            year_close = ws.Cells((i - 1), 6).Value

            'find the yearly change using year_open and year_close, output to the summary table
            yearly_change = year_close - year_open
            ws.Cells(j, 10).Value = yearly_change

            'conditional format for yearly_change
            If yearly_change > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 0
            End If

            'find the percent change and output to the summary table using FormatPercent
            'using if statement to handle any division by zero or no changes
            if year_open <> 0 And yearly_change <> 0 Then
                percent_change = yearly_change / year_open
                ws.Cells(j, 11).Value = FormatPercent(percent_change)
            Else
                ws.Cells(j, 11).Value = "N/A"
            End if

            'output the total_volume to the summary table
            ws.Cells(j, 12).Value = total_volume

            'reset year_open, year_close, and total_volume to 0 for use in next ticker
            year_open = 0
            year_close = 0
            total_volume = 0

            'move to the next row for the summary table
            j = j + 1

        wend

    'if statement will change these variables to go in the right directions
    'set current_great_decrease to first percent
    'set current_decrease_ticker to first percent's ticker
    'set current_great_increase to first percent
    'set current_increase_ticker to first percent's ticker
    
    'set current_great_volume to first total volume
    'set current_volume_ticker to first total volume's ticker
    
    'set table's headers and row labels
    'cells in row 1, columns P and Q = ticker, value
    'cells in rows 2-4, column O = greatest % inc, greatest % dec, greatest total vol
        
    'loop through summary table by row
        'if current row's percent is positive then
            'if current row's percent is greater than current_great_increase then
                'current_great_increase = current row's percent
                'current_increase_ticker =current row's ticker
            'end if
        'elseif current_row's percent is negative then
            'if current row's percent is less than current_great_decrease then
                'current_great_decrease = current row's percent
                'current_decrease_ticker = current row's ticker
            'end if
        '>>> do I need an "else" statement here to account for zero percent?
        'end if
        
        'if current row's total volume is greater than current_great_volume then
            'current_great_volume = current row's total volume
            'current_volume_ticker = current row's ticker
        'end if
        
    'next row
    
    'output current_great_decrease, current_great_increase, and current_great_volume along with their tickers
    'to the new table

    Next ws

End Sub
