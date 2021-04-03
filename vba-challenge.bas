Attribute VB_Name = "Module1"

Public Sub StockMarketData()

 'Declearing variables
 Dim tickerSymbol As String
 Dim yearlyChange As Double
 Dim percentChange As Double
 Dim tickerVolume As Double
 Dim summaryTableRow As Integer
 Dim openPrice As Double
 Dim closePrice As Double
 
 'Assigning initial values
 tickerVolume = 0
 openPrice = Cells(2, 3).Value
 summaryTableRow = 2
 
 'Rule to identify the last row in primary data
 lastrow = Cells(Rows.Count, 1).End(xlUp).Row
 
 'Creating and formatting table headers
 [I1:L1] = Array("Ticker", "Yearly Change", "Percentage Change", "Volume")
 [I1:L1].Interior.ColorIndex = 33
 [I1:L1].Font.Bold = True

 
    'Looping through all rows
    For i = 2 To lastrow
    tickerSymbol = Cells(i, 1).Value 'Ticker Name
    closePrice = Cells(i, 5).Value  'Identifies and picks the closing price for each stock
    yearlyChange = closePrice - openPrice 'Compute the annual price change for each stock
    tickerVolume = tickerVolume + Cells(i, 7).Value 'Computes the total stock volume for the year
    
    'Computing the annual percentage change for each stock
    If openPrice = 0 Then
        percentChange = 0
    Else
        percentChange = (yearlyChange / openPrice) * 100
    End If
    
    
    
    'Checking current cell with the cell below it and updating the summary table as we move through the rows
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Range("I" & summaryTableRow).Value = tickerSymbol
        Range("J" & summaryTableRow).Value = yearlyChange
        Range("K" & summaryTableRow).Value = (Round(percentChange, 2) & "%")
        Range("L" & summaryTableRow).Value = tickerVolume
        
        'Resetting the initial values moving to the next row of the summary table
        tickerVolume = 0
        openPrice = Cells(i + 1, 3).Value
        summaryTableRow = summaryTableRow + 1
    
    End If
    
    Next i
    

   'Below is the condition for highlighting positive change in green and negative change in red
    For i = 2 To lastrow
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        Else
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
 
'Bonus
 'I tried but I was not able to get the code to work as expected. The code below is the closest thing I have
 
 'Declearing variables
 Dim MaxValueTableRow As Double
 Dim maxValue As Double
 Dim minValue As Double
 
 
 'Initializing variables
 MaxValueTableRow = 2
 maxValue = 0
 minValue = 0
 
  
  
  'creating row and column headers for maximum number tables
  [O1:P1] = [{"Ticker", "Value"}]
  [N2] = "Greatest % Increase"
  [N3] = "Greatest % Decrease"
  [N4] = "Greatest % Total Volume"
  
   'checking through the rows of the summary table for max increase, decrease and volume
   'This code only check for the maximum values and I understand that it has a problem, but it is the best
   'I can do, for now.
    For i = 2 To lastrow
        For j = 11 To 12
            If Cells(i, j).Value > maxValue Then
               maxValue = Cells(i, j).Value
                Range("O" & MaxValueTableRow).Value = Cells(i, 9).Value
                Range("P" & MaxValueTableRow).Value = maxValue
            
            
                MaxValueTableRow = MaxValueTableRow + 1
            
            End If
        Next j
    Next i

End Sub

Public Sub RunStockMarketData()

'Declearing variable
 Dim i As Integer
 
 'Setting initial valie
 i = 1
 
 'Running while loop through worksheets
 Do While i <= Worksheets.Count
    Worksheets(i).Select
    
 'Stating which code to run through the loop
    StockMarketData
    i = i + 1
Loop


End Sub
