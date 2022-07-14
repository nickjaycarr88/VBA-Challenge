Sub vbachallenge()
'enabling ws so the script can run accross all pages

For Each ws In Worksheets
'making stock ticker into the string format
Dim stock_ticker As String
'making volume colum as longlong as it is too large to be just a long
Dim volume As LongLong
'making summary table row as integer as it is not a big number
Dim summary_table_Row As Integer
'making last row as long as the amount of rows is too many for integer
Dim LastRow As Long
'making the open price of a stock as double as it is to two decimal places
Dim openprice As Double
'making the open price of a stock as double as it is to two decimal places
Dim closeprice As Double


'starting the summary table at 2 as the first row is for headings
summary_table_Row = 2
'starting the volume at 0
volume = 0
'this is to find out the amount of rows in each worksheet, and to then save it to a variable
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    'the following are the headers of the columns to be inserted
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    'the following formats the columns to expand making the more fitting for the headers and data
    ws.Range("A:L").EntireColumn.AutoFit
'creating a loop to loop through all the rows of each worksheet
For i = 2 To LastRow
    'this if statement compares the value of a cell the cell directly below it. if the are not the same, then proceed to do the folling commands
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'this stores the close price of a stock
        closeprice = ws.Cells(i, 6).Value
        'this stores the ticker symbol for each company
        stock_ticker = ws.Cells(i, 1).Value
        'this gets the volume from each row and adds it to the following volume cells to add up the total volume for each company
        volume = volume + ws.Cells(i, 7).Value
        'this sets the cell location for the stock ticker
        ws.Range("I" & summary_table_Row).Value = stock_ticker
        'this sets the cell location for the volume
        ws.Range("L" & summary_table_Row).Value = volume
        'this sets the cell location for the yearly change
        ws.Range("J" & summary_table_Row).Value = closeprice - openprice
        'this sets the cell location for the percentage change
        ws.Range("K" & summary_table_Row).Value = ((closeprice - openprice) / openprice)
        'this converts the percentage change into a two decimal percentage amount
        ws.Range("K" & summary_table_Row).NumberFormat = "0.00%"
        'this increases the summary table row by one each time it reaches this point
        summary_table_Row = summary_table_Row + 1
        
        
        'this resets the volume to 0 so that the next time it computes the volume for the next company, it starts at 0.
        volume = 0
    'the else thatement if the above if statement is not compatible
    Else
        'this inner if statement allows us to know when the volume is 0. this means that we are iterating over a new company and the first line will be the opening price
        If volume = 0 Then
        'this saves the opening price as a variable
        openprice = ws.Cells(i, 3).Value
        End If
        
       'this saves the volume as a variable
        volume = volume + ws.Cells(i, 7).Value
    End If
Next i
    'this for statement iterates over the values stored in the yearly change column.
    For i = 2 To LastRow
    'if the value of the cell is greater or equal to 0, then turn the cell colour green
        If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        'if the cell value is less than zero, than turn the cell colour red
        Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
    
Next ws
    
End Sub


