Sub stocks_summary():
    dim i as long 'to be used to go through each row in the table
    dim current_ticker as string 'to be used for looping within one ticker at a time and output the ticker to the summary table
    dim year_open as double, year_close as double 'to be used for getting the open and close of the beginning/end of the year
    dim yearly_change as double, percent_change as double 'to be used for the summary table's formulas
    dim total_volume as double 'to be used to add each row's volume for each ticker and getting the total for each ticker
    dim j as long 'to be used to output totals for each ticker into their own rows in the summary table, iterating by 1 for each ticker change
    dim ws as Worksheet 'to use for iterating through each worksheet in the file

    'declaring variables to be used to get values for the second table
    dim increase_number as double, decrease_number as double, volume_number as double
    dim increase_ticker as string, decrease_ticker as string, volume_ticker as string
    dim s as long

    'loop through each worksheet
    For Each ws In Worksheets

        'assign values to variables
        i = 2
        j = 2
        total_volume = 0
        year_open = 0
        year_close = 0
        s = 2


        'set column names for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'set table's headers and row labels
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        

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
                ws.Cells(j, 11).Value = 0
            End if

            'conditional format for percent_change
            If percent_change > 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 4
            ElseIf percent_change < 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(j, 11).Interior.ColorIndex = 0
            End If

            'output the total_volume to the summary table
            ws.Cells(j, 12).Value = total_volume

            'reset year_open, year_close, and total_volume to 0 for use in next ticker
            year_open = 0
            year_close = 0
            total_volume = 0

            'move to the next row for the summary table
            j = j + 1

        wend

    'setting variables for greatest increase, decrease in percentage, and greatest volume
    increase_number = ws.Cells(2, 11).Value
    increase_ticker = ws.Cells(2, 9).Value
    decrease_number = ws.Cells(2, 11).Value
    decrease_ticker = ws.Cells(2, 9).Value
    volume_number = ws.Cells(2, 12).Value
    volume_ticker = ws.Cells(2, 9).Value

    'loop through summary table by row
    While Not IsEmpty(ws.Cells(s, 9).Value)

        'check if percent change is positive or negative
        If ws.Cells(s, 11).Value > 0 Then
            
            'check if value is greater than increase_number
            If ws.Cells(s, 11).Value > increase_number Then
                increase_number = ws.Cells(s, 11).Value
                increase_ticker = ws.Cells(s, 9).Value
            End If
        
        ElseIf ws.Cells(s, 11).Value < 0 Then

            'check if value is less than decrease_number
            If ws.Cells(s, 11).Value < decrease_number Then
                decrease_number = ws.Cells(s, 11).Value
                decrease_ticker = ws.Cells(s, 9).Value
            End If

        End If
        
        'check if total stock volume is greater than volume_number
        If ws.Cells(s, 12).Value > volume_number Then
            volume_number = ws.Cells(s, 12).Value
            volume_ticker = ws.Cells(s, 9).Value
        End If

        s = s + 1
    wend
    
    'output increase_number, decrease_number, and volume_number to table along with their tickers
    ws.Cells(2, 16).Value = increase_ticker
    ws.Cells(2, 17).Value = FormatPercent(increase_number)

    ws.Cells(3, 16).Value = decrease_ticker
    ws.Cells(3, 17).Value = FormatPercent(decrease_number)

    ws.Cells(4, 16).Value = volume_ticker
    ws.Cells(4, 17).Value = volume_number

    Next ws

End Sub
