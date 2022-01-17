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
- Refactoring results in the code being shorter, faster and more efficient. Please see the table below for the time difference comparisons

| | **Original Script** | **Refactored Script** |
| --- | --- | --- |
| 2017 | ![All Stocks (2017)](/Resources/OgTimer_2017.png) | ![All Stocks (2017)](/Resources/RefactorCodeTimer_2017.png) |
| 2018 | ![All Stocks (2017)](/Resources/OgTimer_2018.png) | ![All Stocks (2017)](/Resources/RefactorCodeTimer_2018.png) |

>*Please note these were the times achieved when the code was first run after opening the file. Subsequent runs give different times. To get an overall unbiased contrast the file was re-opened every time to run the different codes for the different years*

### Disadvantages
- The biggest disadvantage to refactoring the code is investing the time in understanding what is already written. Different people think differently and therefore approach problems differently. Due to this, despite providing comments, understanding someone else's chain of thought requires effort

### Pros and Cons 
- Compared to the original code, the refactored script required only 1 *for* loop as opposed to 2 (for the actual analysis, not including the *for* loop for adding data to the new sheet). Which not only reduced the number of times the code was run, but also the confusion that came with multiple loops nested within one another (***Advantage***)

- Trying to understand the approach the new code wanted to take took me some time and effort. Unlike the original code, which had variables that could easily be set to 0 and again assigned a value through the loops, the new code made use of arrays and we were only given a limited number of variables to use. Going from one school of thought to another took time. (***Disadvantage***)