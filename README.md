VBA-challenge

Description:

This code was written to loop through all worksheets of an Excel workbook and take stock market data and output summary data for each ticker symbol. The ticker symbols with the greatest
percent increase, greatest percent decrease, and largest stock volume were also identified.

Note: To run on all worksheets in a workbook, use the loopWsCh2 subroutine. To run on only the active worksheet use the challenge2 subroutine.

References:

This code was written by me, Connor Lorden, with help from the following sources. These sources are referenced in the comments of the code at the locations they were used.

#1	This source was used for statements to format cells with values between 0 and 1 as percentages.
	https://www.statology.org/vba-percentage-format/

#2	This code was used to find the last row of a column. It was used to determine the length of the for loops.
	https://www.excelcampus.com/vba/find-last-row-column-cell/

#3	This source was used for statements to find the maxiumum and minimum values of a column. It was applied in finding the greatest percent increase and decrease.
	https://stackoverflow.com/questions/42633273/finding-max-of-a-column-in-vba

#4	This source was used in writing the loopWsCh2 subroutine. It was applied to run the challenge2 subroutine on all worksheets.
	https://stackoverflow.com/questions/21918166/excel-vba-for-each-worksheet-loop