# Stock-Analysis
Overview: Explain the purpose of the analysis
The purpose of this analysis is to create an easy-to-use tool that will display the total volume and return of various green energy stocks. The intended user is someone with no programming experience and the tool does not require VBA knowledge to utilize. Steve, the indended user, is helping his parents with their investments who have all of their money in a single stock, "DAQO". He has also mentioned using this code in the future to analyze more than only green stocks.


## Results


Initially, the stock DAQO ("DQ") was of highest interest as Steve needs to know if their current investment is worth keeping. However, taking it a step further to analyse all stocks simultaneously would be more advantageous in the end and avoid an unnecesary step. With data on eight stocks for years 2017 and 2018, the following logic flow was implemented: Establish which year is to be analysed, combine all daily volumes for the stock to return yearly volume, determine the return of each stock using the starting and ending values for the given year, and lastly display all data in an easy-to-read table format. 

![Run_All_Stocks_Analysis_Button](https://user-images.githubusercontent.com/95305584/148859105-740d0043-c7b7-46ac-be7e-ba5a8f505d82.PNG)

In order to avoid the need for understanding VBA, a button was created to run the VBA code in the backgroundm(image1). The user is then prompted to insert which year they would like to have analyzed using an input box. Once the program knows which year to choose, it will open the worksheet with the associated year value and use that data. (image of input box and worksheet activate line). The next steps are the key aspects of the code but do not explain every little detail.

![nested_for_loop_image](https://user-images.githubusercontent.com/95305584/148859048-ae18b6e9-00ff-4821-a710-09b7c08a2d3a.png)

By creating an array, each ticker name is stored in alphabetical order starting with "tickers(0)" being equal to "AY" and so forth. This is utilized in the outer loop of the nested for loop that does most of the work. The inner for loop calculates the total volume, starting value price and ending value price of a given stock. The outer loop runs the inner loop for each stock and displays it on the excel worksheet in a table format that highlights each stock's return in eitehr greeen for a postive return or a negative return in red. The last portion formats the data for aesthetic purposes followeed by the message box that displays.

Taking a quick look at the results of 2017 and 2018 stocks, it's clear that 2017 was a significantly better year for the green stocks. The only stock that performed poorly was "TERP" while DAQO had a return of nearly 200%. Nearly the opposite was true when only two of the stocks finished 2018 with a positive return and DAQO dropped back 60%. 

To assess the eficiency of the code, a timer (a feature in VBA) was started at after the year was chosen and stopped upon completion. The timer showed 27.97 seconds when analyzing 2018 data and 29.11 seconds for 2017 data. While this might not seem like much, the data set only looks at 8 stocks and would henceforth take a considerable amount of time when more stocks are taken into account. Luckily, inefficiencies were found and a refactored code was made with significant improvements in run times. The timer now displayed 0.61 seconds for 2018 data analysis and 0.598 seconds for 2017 data.

The changes made were very simple changes to the flow of the code and occured in the for loops. First, the original code shown earlier with the nested for loops indicates the inner loop to start at the beginning of the data and go all the way through to the end looking for the specified stock. That means that it goes through all the data 12 times! To avoid this, the refactored code moves onto the next stock once the previous one is finished being analyzed. This is done by not using a for loop to go through each stock by  making the value in the array change depending on when the inner loop calculates that the last data point for a stock is reached, then addds a value of 1 to go to the next value in thte array. Ovearll, this allows the code to run through the data set only once, saving a lot of procesing time.

![refactored code loops](https://user-images.githubusercontent.com/95305584/148861811-b752ad9c-deda-4776-9b93-c514b6ede989.png)



## Summary: In a summary statement, address the following
1. What are the advantages or disadantages of refractoring code?
  You can run larger amounts of data and not use as much computer memory and won't wait as long. Doesn't search through ALL the data every single time a for loop runs. it idealizes it for a single purpose but copying and pasting for a different data set might be confusing.
  
3. How do these pros and cons apply to refractoring the original VBA script?
 Cons: the stocks listed in the array need to be in the order that they appear in the data set and the data set has to be filtered ahead of time so that all rows of data of a single stock ticker is together with no other stocks in the middle. 
