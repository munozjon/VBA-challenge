

'declare i as an integer, set equal to 2 (the first row of data in the sheets)
'declare current_ticker as string, to be used for 
'declare year_open, year_close as double, to be used for getting the open and close of the beginning/end of the year

'while current cell is not empty (i, 1)
    'current_ticker is set to the current row's first column value AKA the ticker
    'current_ticker = Cells(i, 1).Value
    'while the current cell(i, 1).value equals the current_ticker
        'setting year_open to the first row's "open" by using if statement to check if it's empty
        'this will prevent the variable from being overwrriten within the loop's iterations
        'if year_open is empty Then 
            'year_open = cell(i, 3).value
        'End if

        ' move on to the next row by adding 1 to i
        'i = i + 1

    'wend
    'setting year_close = cell((i - 1), 6).value
    '(i - 1) used to get the last row of the previous ticker

    '____more code_____

    'setting year_open value = 0 to reset the variable for the next ticker
    'setting year_close value = 0 to reset as well
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