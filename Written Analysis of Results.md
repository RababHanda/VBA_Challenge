# VBA of Wall Street

## Overview of Project
The project analyzes the shares of 12 different stocks being traded in the years 2017 and 2018. 

### Purpose
The purpose of this project is to code in VBA. Majority of the tasks revolve around using loops and conditionals to achieve the results requested.  

The project specifically looks at the total volume of shares traded for each stock along with the annual returns. The project then compares the trading between two different years.


## Results and Comparisons
| **Performance in 2017** | **Performance in 2018** |
| --- | --- |
| ![All Stocks (2017)](/Resources/AllStocks_2017.png) | ![All Stocks (2017)](/Resources/AllStocks_2018.png) |

As can be clearly seen in the tables above, overall stocks performed better in the year 2017 than in 2018.

1. The largest contrast can be seen for **DQ**. It was the highest performer of 2017 with an annual return of 199.4% but became the lowest performer of 2018 with a return of -62.6% 

2. The only stock that showed significant improvement from 2017 to 2018 was **RUN**, going from 5.5% to 84.0%

3. **ENPH** showed the most stable performce, albeit a decrease. It went from 129.5% in 2017 to 81.9% in 2018


## Summary
This project required refactoring a pre-written code

### Advantages
- Refactoring results in the code being shorter, faster and more efficient

| | **Original Script** | **Refactored Script** |
| --- | --- | --- |
| 2017 | ![All Stocks (2017)](/Resources/OgCodeTimer_2017.png) [^1] | ![All Stocks (2017)](/Resources/RefactorCodeTimer_2017.png) [^1] |
| 2018 | ![All Stocks (2017)](/Resources/OgTimer_2018.png) [^1] | ![All Stocks (2017)](/Resources/RefactorCodeTimer_2018.png) [^1] |
[^1]: Please note these were the time achieved when the code was first run after opening the file. Subsequent runs give different times. To get an overall unbiased contrast the file was re-opened every time to run the different codes for the different years

### Disadvantages
-