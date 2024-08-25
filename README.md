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

- For Each ws In ThisWorkbook.Worksheets
        Set dataRange = ws.Range("K1:K" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
- maxTicker = cell.Offset(0, -2).Value 'Register the value of 2 cells to the left of MaxVal cell.
