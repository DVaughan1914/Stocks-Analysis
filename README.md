# Stocks-Analysis

# Overview of Project:
Steve, my client, loves the workbook I prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although the code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.


# The purpose and background are well defined (2 pt).
I prepared a workbook for my client Steve and it seems that he is quite pleased with the results.Steve loves the workbook you prepared for him.Created a interactive button within the workbook for him to run the analysis on the stocks he provided at the click of that button.In his efforts to provide his parents with the best options for green stock as possible he wanted to expand the stock analysis. Although my code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute. So I refactoted the codee to run more efficiently while running for more stocks. I ran the analysis of the stocks with a timer embeded to give a elapse time on how lond does it take to run the tool for 2017 and 2018.

# Results
In running the  refactored code for both years (2017 and 2018), there were some improvent in the perfromance of the code when compared to the orignal/inital code before refactoring. The unit of measurement is based on time (seconds or fraction of seconds). 

# The analysis is well described with screenshots and code (4 pt).
In analyzing the results of the refactored code compared to the original code, we see significant improvement in  the time that it takes to run for each year (2017 and 2018).
In the orignal code ( not refactored) we see that it took a time of 1.18 seconds for the code to run.
![image_initial_run](https://user-images.githubusercontent.com/100040621/158076961-7a490ab5-cb16-4c9e-ad19-ccb74aca6011.png)
Then with the refactored code we run it again on the stocks for the year 2017, we was that the time it took to run actually improved from the orinigal run time. Then code ran in 0.57 seconds. Which is almost a whole half of a second improvemenmt forom the original as seen here:
![VBA_Challenge_2017](https://user-images.githubusercontent.com/100040621/158077167-98a825e0-a317-4d06-afa2-7c71a8c6cbff.png)
The results with running the 2018 stocks were even faster than running the 2017 stocks. The 2018 stocks took 0.21 seconds to run the refactored code and to give us the results. We can see from the image below:
![VBA_Challenge_2018](https://user-images.githubusercontent.com/100040621/158077264-ab66385e-fcbc-44ec-bab3-168d543f13be.png)
One may wonder why the 2018 stocks ran faster than the 2017 stocks. That has to do with the stock's activity during each particualr year. Whether the stock had alot oof changes, going up or down, how often and for how long.
The code from the original to the refactored is significantly different, these differences were the reason why the code ran faster with more detailed results. For example, in the refactored code we set an array of key stocks that fit our criteria as seen here
 'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
 So this helps the code focus on these stocks, also because we focused in on a particular time frame that definitely helps. Creating our own code to track the stock ticker within the analysis also played a key role. I made the ticker create an output for us to see with this code:
 
  '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
 Comparing starting prices/ending prices and volume of stock is trading in each year.
 
# Summary
Refractoring the code definitely improved the analysis and the code's performance between the two years in comparison to the oringal code and analysis. I was able to do more in less time. The refactored code is signifianctly longer in lenght and its is more complex but the results speaks for itslef. Refactoring a code for analysis should be very common in any analysis which includes writing code. We will find new and more effective ways to achieve the goal. Codes will evolve from its original stated into something more complex and some times more effective. As we learn and do more research we might learn new and effective ways to complete a task. So going back and revising a code or process to make it more effective, ie. better is rarely a negative. So the customer was even more pleased wit the results and with the time it took.

# There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
  # Advantages
The purposes of refactoring according to Martin Fowler (Father of Code Smell) are stated in the following:

Refactoring Improves the Design of Software
Refactoring Makes Software Easier to Understand
Refactoring Helps Finding Bugs
Refactoring Helps Programming Faster
  # Disadvantages
Below dangers can happen after or during the code refactoring,
It is expensive and risky in the view of management.
It may introduce bugs.
Delivery schedule is very tight.
Management doesn't care about maintainability and extension of code base.

# There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
  # Original script

The original code was shorter and takes a sorted time to write and read, its simple in its construction.
The original code was not effecitent and didnt provide the use with detailed results, takes longer to run.
  # Refactored Script
The refactored code was more efficent on time, it provided more detailed information, best used to make decisions on the given problem or question.
The refactored code is longer in length, it takes more time to write and test, can introduce bugs and problems. Debugging can be timely and could cause one not to meet required deadline.
