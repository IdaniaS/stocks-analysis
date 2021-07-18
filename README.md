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

#### Comments
When refactoring code, it is not adding new functionality. The code is made more efficient-by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. The comments and documentation describing sections that are edited help other readers with following the code. This means avoiding adding comments that are not helpful or obvious to the reader are best described as being a Dry addtion so you are not repeating information. The best way to keep the code looking clean and easy to read comments is through having a consistent indentaion and code grouping. These comments can be viewed in the photo additions above as a short explanation to what the code is preforming.


![page4](https://user-images.githubusercontent.com/86635590/126084508-55d818fd-e745-4911-89b8-9fb6c47baece.PNG)
![page50](https://user-images.githubusercontent.com/86635590/126084553-c6de3799-e5c5-47c6-8387-adf7ce94ca48.png)

The outputs in the 2017 and 2018 stock analysis from the file VBA_Challenge.xlsm matched the AllStockAnalysis in the module. Which then allowed the dataset to be anaylized and run through the refactored code.

### Dataset 2017: VBA_Challenge_2017.png
![2017Stocks](https://user-images.githubusercontent.com/86635590/126084754-89cdbc91-bd5c-429b-a63b-f88aca7ae11f.PNG)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/86635590/126085225-0261dd05-a731-40f6-b7c6-8eb9bad3a718.PNG)

### Dataset 2018: VBA_Challenge_2018.png
![2018Stocks](https://user-images.githubusercontent.com/86635590/126084763-e69d5979-b060-47fa-a406-cc773b80c31b.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/86635590/126085234-779c4cd0-a11e-4c38-ae47-6364addbff3b.PNG)


## Summary
### Advantages of refactoring code
- Structure of the code appears well written which allows for the debugging of errors to be approached in a simpler manner
- Simplified code has a possibility of saving time and money in the future as it simplifies the implementation
- The code makes it easier to reveal patters in a more comprehensible manner becasue it is not as messy
### Advantages and Disadvantages of the original and refactored script
- The procedure in refactoring can be expensice and risky to implement which mean the original has the advantage of cheaper and holds less risk when running the code
- Time and resources will be dedicated to cleaning up the code which may introduce bugs or very tight scedhule delivery deadlines
- Complexity of unstructured code is usually best to split in several functions but harder to determine where sections need to be added or edited

Code refactoring is the process of restructuring an existing code without changing its external behavior which is intended to improve the implementation. The refactoring is alwasys something to consider when writting/editing codes. It has the advantages of being viewed quickly and saving both time and money. However, it can also be tough when dealing with duplications within the structure and tight delivery deadlines. It is a situational case but the clean and well-organized code will produce a better faster outcome in the long run.
