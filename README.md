# VBA-challenge
This is a VBA Challenge Project
This project recaps the 4 components of programming: variables, iterations, functions, and conditionals.
The program was divided into separate functions to ensure it ran more efficiently and the debugging was faster.  
- Stock(): Adds 4 new columns to each worksheet (Ticker, Quarterly Change, Percent Change, and Totatl Stock Volume), and calculates the total values for each ticker placing the values in the table.
- FindMaxValueAndTicker(): It creates a table to hold the Maximum percent increase and its ticker and finds the mentioned data.
- FindMinimumValueAndTicker(): Function that finds the minimum percent change and its correspondent ticker
- GreatestTotalVolume(): Finds the Maximum volume and its ticker.
- FormatColumns(): made formatting changes to some columns, like adding %, $, color, etc.

Besides the topics learned during class, I used Google and Xpert Learning for help. Most of the help was for debugging. I also used The following snippets from Xpert Learning Assistant:

- For Each ws In ThisWorkbook.Worksheets 'Expression for class was giving me an error message
- Set dataRange = ws.Range("K1:K" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)  'Expression for class was giving me an error message
- For Each cell In dataRange 'Cell declared as Range
- If IsNumeric(cell.Value) Then 'Checks if the value is a number
- maxTicker = cell.Offset(0, -1).Value 'Looks for the value of the cell located 1 column to the left
- ws.Columns("B:B").NumberFormat = "mm/dd/yyyy"  'Format date
- ws.Columns("J:J").NumberFormat = "$#,##0.00"   'Format currency     
- ws.Columns("K:K").NumberFormat = "0.00%"       'Format percentage 
