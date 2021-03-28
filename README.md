# stock-analysis
## Overview of Project:
The purpose of this analysis was to help Steve with his stock analysis by using VBA and Excel to make it easier for both him and his parents to use. Steve wanted to know about how the 2018 stocks performed. We had to measure the yearly return and compare it to the previous year, but he also wanted to analyze other stocks since DAQO might not be the best stock for his parents to choose.

## Results:
The stock performance between the 2017 and 2018 stocks shows the 2017 stocks performed better than 2018. 2017 stocks had a better return overall compared to 2018 since 2017 only had one stock in the negatives for their return, which was TERP, while 2018 had ten out of twelve stocks in the negatives for their return. The only two stocks that made a positive return in 2018 was ENPH and RUN compared to the 11 positive returns in 2017. This shows that if Steve were to get his parents to choose a stock, then ENPH and RUN would be a better option for them to choose because those two stocks stayed consistent with a positive return in both 2017 and 2018. 

<img width="215" alt="All_Stocks_VBA_Challenge_2017" src="https://user-images.githubusercontent.com/79742633/112744182-98e27400-8f52-11eb-8585-39c4ecb9705e.png">
<img width="216" alt="All_Stocks_VBA_Challenge_2018" src="https://user-images.githubusercontent.com/79742633/112744191-aa2b8080-8f52-11eb-876a-4b058b306b0b.png">

**Original Script**

The execution times of the original script was longer than the refactored script. The original script time took 0.84375 seconds for the 2017 stock analysis and the original script for the 2018 stock analysis took 0.796875 seconds to finish running. 

<img width="235" alt="Original Code Time 2017" src="https://user-images.githubusercontent.com/79742633/112744195-be6f7d80-8f52-11eb-965b-01a2eccf219b.png">
<img width="236" alt="Original Code Time 2018" src="https://user-images.githubusercontent.com/79742633/112744210-d0512080-8f52-11eb-9cff-9f25d8430a66.png">

**Refactored Script**

My refactored script took 0.1640625 seconds each for both the 2017 and 2018 refactored analysis to show the results, which is still faster than the original script.

<img width="240" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/79742633/112719913-af3bf180-8eb8-11eb-8729-7b2080cf65c5.png">
<img width="247" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/79742633/112719972-0fcb2e80-8eb9-11eb-965a-e994ae57040f.png">


## Summary:
**What are the advantages or disadvantages of refactoring code?**

Advantages of refactoring code could help make your code look more organized and easier to read and understand. For example, your code might not have indentations and comments at first, but then you added it after to make it easier for others to understand. Another advantage would be that is can save a few seconds of time to analyze data.
Disadvantage would be accidentally messing up your code and then having it not work the way you want it to, which that may be difficult to fix if you didn’t save a copy of it. Sometimes there are more than one way to do a specific task, but it may not work if you write the new code wrong.

**How do these pros and cons apply to refactoring the original VBA script?**

These pros apply to my original VBA script because it helped put together the information for the year I searched run faster. It would’ve been very time consuming if I manually did that without VBA. Cons would be it can mess up your code and give you a difficult time to solve the problem and get it working again.
