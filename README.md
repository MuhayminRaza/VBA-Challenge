# VBA-Challenge
The Module 2 challenge required a code such that it looped through all the stocks in the given years of 2018-2020 amd output the following information: 
- The ticker symbol
- Yearly change from the opening price 
- The percentage change from the opening price
- The total stock volume of the stock

The code also found the stocks with the 
- greatest percentage increase 
- greatest percentage decrease
- greatest total volume 

The files in the "Screenshots" folder show the results of the code for 2018-2020 respectively. 

Code Sources:
The following excerpts from the code script were completed with the help of a TA in office hours:

        'Set Ticker Counter to first row
        TickCount = 2
        
        'Set start row to 2
        j = 2
        
        'Find the last non-blank cell in column A
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox ("Last row in column A is " & LastRowA)

        'Find last non-blank cell in column I
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox ("Last row in column I is " & LastRowI)


        'For greatest decrease--check if next value is smaller--if yes take over a new value and populate ws.Cells
         If ws.Cells(i, 11).Value < GreatDecr Then
         GreatDecr = ws.Cells(i, 11).Value
         ws.Cells(3, 16).Value = ws.Cells(i, 9).Value

            
        'Djust column width automatically
        Worksheets(WorksheetName).Columns("A:Z").AutoFit

