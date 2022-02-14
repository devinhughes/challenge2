## challenge2
# Stock Analysis
This project will consist of us refactoring VBA code to loop through Stock results more efficiently.
# Overview of Project
In this project Steve had asked for help analyzing different stocks so that his parents can make the best decision about what stock to invest in.
# Purpose
The purpose is to refactor the VBA code so that we can loop through only once and having a more efficient run time.
# Results
To refactor the code we created output arrays for each of the variables. The goal is to make our analysis more efficient by reducing the number of times the code must loop through different variables. We will keep the array from the original code for tickers. We created the output arrays tickerVolumes(), tickerStartingPrices(), and tickerEndingPrices(). The created variable tickerIndex is used to refer to the tickers in our initial array. We set this variable equal to zero and prepared to run the loop through the datasheet.
![image](https://user-images.githubusercontent.com/89948353/153945719-a06d88cb-7fd5-436c-9573-b74547a90046.png)
To run an analysis through the spreadsheet we created a loop to run through the dataset. Just like the original code had done. We first created a ticker to increase volume as the code ran through each cell. The variable tickerVolumes(i) was set equal to zero in a previous For Loop. In the current For loop it was run through each cell and added each volume until we reached the end of the ticker by referencing tickerIndex as seen in the code below.
![image](https://user-images.githubusercontent.com/89948353/153945741-19cf6b1d-f018-4ce9-8d81-9f75d4927b4c.png)
What was then needed was to find our ending prices and starting prices. Instead of creating a nested For loop and running a new loop for each ticker to find volume and then looping through rows to find starting/ending prices, what can be used is our variable tickerIndex to loop through more efficiently.
![image](https://user-images.githubusercontent.com/89948353/153945786-2caa590c-4957-4214-adda-ea51515f823e.png)
For the final part of the project we created a new datasheet and were able to display the results for each variable using our arrays for each ticker. When the code was run, the refactored code saw an increased performance for both data sets. The original codes ran for .496 seconds for 2017 and .484 seconds in 2018. The new times are shown below.
![image](https://user-images.githubusercontent.com/89948353/153945828-f1b343f8-d4ec-4cdd-a481-d1e9f7dd5f5d.png)

![image](https://user-images.githubusercontent.com/89948353/153945840-a352361b-b2e6-4277-98c0-3fd6199e9017.png)

![image](https://user-images.githubusercontent.com/89948353/153945852-c237a284-dd3a-4574-8054-c2d0bdb4fd6e.png)

# Advantages and Disadvantages of Refactoring Code
Refactoring has great advantages as it allows the collaborative nature can bring a new take and add different ways of coding. There is always room for improvement, and the collaboration allows coders a chance to improve their code. It can help to make each code made more effective and improve the readability of the code. Refactoring code can be a disadvantage if is not effective. This could be if the code used is out of date and other methods and programs can work better than the one used. It might not be as fruitful as just starting over in some cases.
# Advantages and Disadvantages of Refactored VBA script
There are many advantages of refactored VBA script. A large advantage is how efficient refactoring can be by comparing coding time. As seen in our results of our refactored code, we have significantly improved out code performance. The original codes ran for .496 seconds for 2017 and .484 seconds in 2018. The performance times of .199 for our 2017 analysis and .191 time for our 2018 analysis. This project shows just how effective refactoring code can be in VBA script. A main disadvantage is that if there is a language or program that works better for the out project we would not be able to refactor the VBA script but would have a better time restarting the code in a new program just like we see in the coding.

