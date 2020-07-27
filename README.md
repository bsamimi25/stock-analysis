# Stock Analysis Challenge #2
## Overview of Project
The purpose of this project was to refactor the code used during this weeks module. We wanted to increase the timing of our code to be able to process more data in the future with more tickers on the excel sheet, hundreds more tickers to be exact!
## Results
After I refactored the code, my run was indeed faster and will be able to hold thousands amounts of data points for future use for our friend Steve here. 
Now comparing the code before the refactoring we can definitely see how there are pro and cons to this new code. 
Getting into more of the details, we see that in our refactored code we use this tickerIndex variable which will take up the space of each ticker, for tickerVolumes, tickerStartingPrices and tickerEndignPrices. 
I started by setting the tickerIndex to 0, and only using one loop to go thru all of the rows and collect the data. I did not use any nested loops in this code which resulted in a faster run time. 
Now my tickerVolumes, tickerStartingPrices, and tickerEndingPrices all had the tickerIndex used to get their data points for each ticker. Example being tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value. 
The most crucial part of the code is this:   
If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                tickerIndex = tickerIndex + 1
The tickerIndex= tickerIndex + 1 line provides the ticker change for the loop. For example, tickerIndex= 0+1 for the first time of code when we stated at the start that tickerIndex was equal to 0 then it becomes the ticker CSIQ, since ticker(1)="CSIQ".
Now lets compare our actual refactored and original code. 
![](https://github.com/bsamimi25/stock-analysis/blob/master/VBA_Challenge_2017Refactored.png)
![](https://github.com/bsamimi25/stock-analysis/blob/master/VBA_challenge_2017.png)
![](https://github.com/bsamimi25/stock-analysis/blob/master/VBA_Challenge_2018refactored.png)
![](https://github.com/bsamimi25/stock-analysis/blob/master/VBA_Challenge_2018.png)
Here as stated before we see the big difference betweeen the refractored and original code in terms of time to output the data. The refactored time is shorter than the orignal code and evidence is in the imges provided. 
## Summary 
Lets dive into the pros and cons of this refactored vs nested loop code. 
To start, the pros are easy with that the refactored code runs faster and can withstand more tickers/data points which the lack of nested loops within it. The one loop allows for a quicker analysis because of the use of the tickerIndex which becomes an index to the ticker arrays tickerVolumes(12), tickerStartingPrices(12) and tickerEndingPrices(12). The tickerIndex is getting incremented with each new ticker, not by going thru a loop.
The original code did not Index the volumes, starting, and ending prices. 
Cons to the refactored code in my honestly opinion was trying to figure out why we needed to set tickerIndex to zero at the start. It completely threw me off of how and why we started at zero. The code runs very fast and is indeed an improvement. In my opinion I couldn't see a con other than there could have been a better way to create the tickerIndex but the code had a significant shorter run time compared the original. 
The original code itself is a con in that it can not run as fast and will take a longer time in the future when there is more amounts of data to loop through.  
Here we can see how the refractored loop in the long term is better. To start this is the orginal code where we see we do in fact have nested loops to loop through all the rows and gather the designated data. The nested loops are in fact a disadanvatage here because it takes more time to go through the rows.  

    For i = 0 To 11
      ticker = tickers(i)
      totalVolume = 0
    
    '5) loop thru rows in the data
    Worksheets(yearValue).Activate
    For j = 2 To RowCount
        
        '5a) Get total volume for current ticker
         If Cells(j, 1).Value = ticker Then
            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(j, 8).Value
        
            End If
        
        '5b) get starting price for current ticker
         If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value
        
            End If
              
        '5c) get ending price for current ticker
          If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value
        
            End If
    
    Next j
    Next i 

   We are seeing that there is no nested loops in our refactored model.  Also can see where the tickerIndex is used to go thru all the arrays to extrct each ticker data. Now the one loop allows us to loop through the data all at once compared to the nested loops in the original code. This one loop notion leads to more of an advantage especially when there is more data involved:
    
   
     For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
   
    '6b) loop over all the rows
    For i = 2 To RowCount
            
        '7a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
      
        '7b) Check if the current row is the first row with the selected tickerIndex.
             If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                 tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
        '7c) check if the current row is the last row with the selected ticker
        '7d Increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                tickerIndex = tickerIndex + 1
                 
            End If
            
    Next i

In conclusion, our refactored model is more efficient in its analysis process. Using the tickerIndex we are seeing how it is being used as an index to all three arrays with the combination of not having a nested loop leads to a faster processing time. 
