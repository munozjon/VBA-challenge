

'dim i as an integer, set equal to 2 (the first row of data in the sheets)
'dimlare current_ticker as string, to be used for looping within one ticker at a time
'dim year_open, year_close as double, to be used for getting the open and close of the beginning/end of the year
'dim total_volume as integer, set equal to 0, to be used to add each row's volume for each ticker and getting the total for each ticker
'dim j as integer, set equal to 2, to be used to output totals for each ticker into their own rows in the summary table, iterating by 1 for each ticker change

'set column names for the summary table
'Cells(1, 9).Value = "Ticker"
'Cells(1, 10).Value = "Yearly Change"
'Cells(1, 11).Value = "Percent Change"
'Cells(1, 12).Value = "Total Stock Volume"



'while current cell is not empty (i, 1)
    'current_ticker = Cells(i, 1).Value, to set the variable to the current row's ticker
    'while the current cell(i, 1).value equals the current_ticker
        'setting year_open to the first row's "open" by using if statement to check if it's empty
        'this will prevent the variable from being overwrriten within the loop's iterations
        'if year_open is empty Then 
            'year_open = cell(i, 3).value
        'End if

        'add row's volume to the "total_volume" variable
        'total_volume = total_volume + Cells(i, 7).Value
        
        ' move on to the next row by adding 1 to i
        'i = i + 1

    'wend
    'setting year_close = cell((i - 1), 6).value
    '(i - 1) used to get the last row of the previous ticker

    '____more code_____

    'setting year_open value = 0 to reset the variable for the next ticker
    'setting year_close value = 0 to reset as well
    'j = j + 1, to move to the next row for the summary table
'wend





Sub stocks_summary():
    Dim i As Integer
    i = 2
    Dim current_ticker as String
    current_ticker = ""
  
    While Not IsEmpty(Cells(i, 1))
        current_ticker = Cells(i, 1).Value
        While Cells(i, 1) = current_ticker
            'some code
            i = i + 1
        Wend
        'some code
    Wend
    'some code
End Sub



'iterating through multiple worksheets:
    'For each ws in Worksheets