Attribute VB_Name = "Module11"
'All references in comments are in the README of this repo

'Initiate subroutine to loop through all worksheets in the active workbook
'RUN THIS SUBROUTINE
Sub loopWsCh2():

'Declare a variable for the active worksheet
Dim ws As Worksheet

'This wiil move Excel into the worksheet of the current count and run
'the challenge2 subroutine and then move to the next worksheet and repeat for all sheets
For Each ws In Worksheets  'Reference #4
    ws.Activate            'Reference #4
    Call challenge2        'Reference #4
Next ws                    'Reference #4

'End of this subroutine
End Sub

'Initate subroutine to make neccessary edits to individual worksheet. This is what is called upon in the loopWsCh2 subroutine
Sub challenge2():

'Declare variables for later use
Dim outputRow As Integer
Dim volume As Double
Dim op As Double

'Insert text into cells that will act as labels for the outputs
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'create variable to find last row of column 1 and initiate values for outputRow, volume, and op to be used in loops
lastRow = Cells(Rows.Count, 1).End(xlUp).Row 'Reference #2
outputRow = 2
volume = 0
op = Cells(2, 3).Value


'Initiate loop. Given that the ticker value in the first cloumn is the same as the next row, this loop will add the stock volumes
'and store them as a variable. When the ticker value in the next row of the loop changes, it will: output the volume, calculate and
'output the annual change in stock value, format those cells, calculate and output the percent change, and then reset all variables
'for use in calculating the next tickers values.
For i = 2 To lastRow
    If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
        volume = volume + Cells(i, 7).Value
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Cells(outputRow, 9) = Cells(i, 1).Value
        Cells(outputRow, 10) = -(op - Cells(i, 6).Value) 'Added a negative sign to correct yearly change results
            If Cells(outputRow, 10) > 0 Then
                Cells(outputRow, 10).Interior.ColorIndex = 4
            Else: Cells(outputRow, 10).Interior.ColorIndex = 3
            End If
        Cells(outputRow, 11).Value = -((op - Cells(i, 6).Value) / op) ' Added a negative sign to correct percent change results
        Cells(outputRow, 11).NumberFormat = "0.00%" 'Reference #1
        op = Cells(i + 1, 3)
        volume = volume + Cells(i, 7).Value
        Cells(outputRow, 12) = volume
        volume = 0
        outputRow = outputRow + 1
    End If
Next i
   
'Define varaibles for table for second table and define variable to find the last row of the table generated by the previous for loop
lastrow2 = Cells(Rows.Count, 11).End(xlUp).Row 'Reference #2
maxi = Application.WorksheetFunction.Max(Range("K:K")) 'Reference #3
mini = Application.WorksheetFunction.Min(Range("K:K")) 'Reference #3
stock = Application.WorksheetFunction.Max(Range("L:L")) 'Reference #3
   
'Initate loop for second table. This loop will find the ticker value for the ticker with the largest stock volume, maximum percent change,
'and minimum percent change
For j = 2 To lastrow2
    
    If Cells(j, 12).Value = stock Then
        tickstock = Cells(j, 9).Value
    End If
        
    If Cells(j, 11).Value = maxi Then
        tickmax = Cells(j, 9).Value
            
    ElseIf Cells(j, 11).Value = mini Then
        tickmin = Cells(j, 9).Value
        
    End If
Next j
  
'This popultes the maximum and minimum percent changes, the largest stock volume, and their respective ticker values in a table
Cells(2, 16).Value = tickmax
Cells(2, 17).Value = maxi
Cells(2, 17).NumberFormat = "0.00%" 'Reference #1
Cells(3, 16).Value = tickmin
Cells(3, 17).Value = mini
Cells(3, 17).NumberFormat = "0.00%" 'Reference #1
Cells(4, 16).Value = tickstock
Cells(4, 17).Value = stock


End Sub
