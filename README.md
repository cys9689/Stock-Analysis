# Stock Analysis 
## Language Usage: VBA
This project is a VBA practice. The AllStock Analysis sheet is a template that is capable of calculating the yearly return for stock in 2017 or 2018.
The Refactor sheet optimize the template code. Instead of loop through all the stock data for each ticker we analyzed, the refactor code loop through the data only once and collect all of the information it needs in a single pass. 


## Refactor code
The Refactor code has same function as the original. However, it has better efficiency while running. The original code can only be capabable with small dataset, when the number of stock increase to hundred or even thousands, it will be slow. 
'''
tickerIndex = 0
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    For i = 2 To RowCount
    
        'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
            
        End If
  '''
  
  We create a tickerIndex start from zero checking and utlizing the fact that the ticker was in alphabet order. We loop through the column once and check whether the ticker name of the column A fitting the tickers(tickerIndex) or not. if not the starting price will be the cells value. Similar with the starting price, the ending price use the same properties. 
