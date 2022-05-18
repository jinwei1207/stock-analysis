# Stock Analysis With Excel VBA
Click here to view the Excel file: [VBA Challenge - Stock Analysis](https://github.com/jinwei1207/stock-analysis)

## Overview of Project
### Purpose
The purpose of this project is to find the performance on different stocks in  2017 and 2018 
### The Data
The data includes two charts with stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price,  and the volume of the stock. 

## Results
### Analysis

    
    1a) Create a ticker Index
    tickerindex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12)  As Single
    Dim tickerEndingPrices(12)  As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerindex) Then
            tickerStartingPrices(tickerindex) = Cells(i, 6).Value
            
        End If
          

           
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
         If Cells(i + 1, 1).Value <> tickers(tickerindex) Then
            tickerEndingPrices(tickerindex) = Cells(i, 6).Value
            
       
        
               
               
          
            '3d Increase the tickerIndex.
        
          tickerindex = tickerindex + 1
           
        End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
    Next i
   
   ![image](https://user-images.githubusercontent.com/104603177/168940577-039c5e2e-cf1c-4fd6-96b7-7c105a803b04.png)
![image](https://user-images.githubusercontent.com/104603177/168940603-e1426d93-e1fe-4836-96bb-df3d45a76de7.png)

   




## Summary


### performance of the different stocks
Obviously in 2018, most of the stocks are not performed good compared with 2017 but there is two with code ENPH and RUN it increase rapidly. That is how macroeconomic affected the market in 2018
### Pros and Cons of Refactoring Code
Refactoring code helps make the date visualized. Person who can easily find the preformance by given instructions .In reality, huge data it is super hard to be visualized by using this code

### The Advantages of Refactoring Stock Analysis

The biggiest of the Refactoring Stock VBA is that code is running less time than original one.



