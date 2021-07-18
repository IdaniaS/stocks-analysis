# stocks-analysis
## Overview
### Background
Steve loves the workbook prepared for him. At the click of a button, 
he can analyze an entire dataset. 
Now, to do a little more research for his parents, 
he wants to expand the dataset to include the entire stock market over the last few years. 
Although the code works well for a dozen stocks, 
it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

### Purpose
The purpost of this project is to determine whether refactoring a code will successfully make the VBA script run faster.
When refactoring code, you aren’t adding new functionality; you just want to make the code more 
efficient—by taking fewer steps, using less memory, or 
improving the logic of the code to make it easier for future users to read. 

## Results
### Analysis
To use the starter code provided in this Project to refactor the VBA Script dataset so as to loop through the data one time and collect all of the information. Analyzing that the refactoring code successfully made the VBA script run faster. 

Create a tickerIndex variable and set it equal to zero before iterating over all the rows. This tickerIndex will be used to access the correct index across the four different arrays. The four different arrays include the tickers array, tickerVolume, tickerStartingPrices, and tickerEndingPrices arrays.

![page10](https://user-images.githubusercontent.com/86635590/126084274-fd6e9f79-4e33-498a-95f1-8bf5d0994a02.png)
![page20](https://user-images.githubusercontent.com/86635590/126084322-129b5449-6256-4a84-9ac0-8324a5e97a80.png)

A loop was created to initialize the tickerVolumes to zero. Then a for loop that would loop over all the rows in the spreadsheet. The tickerIndex vairable written a script that increases the current tickerVolumes variable that would add the tickerVolume for the current stock ticker.

![page30](https://user-images.githubusercontent.com/86635590/126084422-db27dee1-4a6d-4eea-863b-a167bada4daf.png)

The script loops through all the stock data, reading the values from each row. The if-then statment is used to check the current row selected tickerIndex. If the statement matches the variable then it continues to produce the prefered outcome. One if-then statement was focused on formatting the cells in the spreadsheet based on returned valuse. If the data returned was positive the cell was formatted green and if the outcome was negative the cell turned red. 

![page4](https://user-images.githubusercontent.com/86635590/126084508-55d818fd-e745-4911-89b8-9fb6c47baece.PNG)
![page5](https://user-images.githubusercontent.com/86635590/126084525-8b3fdfe0-7aed-484d-bf96-bfa80aeb7dfa.PNG)



## Summary
### Advantages and Disadvantages of refactoring code
### Advantages and Disadvantages of original code
