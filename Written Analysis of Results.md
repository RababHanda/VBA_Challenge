# VBA of Wall Street

## Overview of Project
The project analyzes the shares of 12 different stocks being traded in the years 2017 and 2018. 

### Purpose
The purpose of this project is to code in VBA. Majority of the tasks revolve around using loops and conditionals to achieve the results requested.  

The project specifically looks at the total volume of shares traded for each stock along with the annual returns. The project then compares the trading between two different years.

## Results and Comparisons
| Performance in 2017 | Performance in 2018 |
| --- | --- |
| ![All Stocks (2017)](/Resources/AllStocks_2017.png) | ![All Stocks (2017)](/Resources/AllStocks_2018.png) |

### Analysis of Outcomes Based on Launch Date

**This analysis examines the campaigns for Theater**

1. A new column to extract years from the entire launch date was created in the KickStarter sheet. The dates were already in **short date** format, so it was simple to use the **Year()** function by just adding the cell reference between the parantheses. This additiona column was added to make the filtering in the pivot table easy.

2. Next, a Pivot Table was created to filter all campaigns according the the *Parent Category* and the *Year* of the launch date. The *Outcomes* were added to the columns so the number of failed, canceled and successful could be easily reflected next to the date it launched on, therefore "Launch Dates" were input in the rows tab. Lastly the "Count of Outcomes" was added to the values tab to consolidate each outcome. Note that only months were kept in rows to classify launch dates into larger groups and see a higher level trend. The parent category was set to Theater to only analyze the trends for that media category.

3. Lastly, a line pivot chart with markers was created to represent the three different outcomes of the campaigns visually. Lines charts are best to represent a trendline over a span of time. The following chart shows how many of each outcomes occured from January to December across all years. 

![Theater Outcomes by Launch Dates ](/Resources/Theater_Outcomes_vs_Launch.png)

### Analysis of Outcomes Based on Goals

**This analysis examines the campaigns for Plays**

1. A new column was created in a new sheet within the same workbook with the fundraising goals grouped in increments of $5000 to combine extrememly varying sums of money into like categories. The outcomes were also regrouped as *Number Successful*, *Number Failed* and *Number Canceled*. The following function was used, with changes made to the criteria depending on what was being input into the columns: 
`=COUNTIFS(KickStarter!$F:$F, "SUCCESSFUL", KickStarter!$D:$D, "<1000",KickStarter!$R:$R,"PLAYS")`

2. Next, a *Total Projects* columns was populated for each fundraising goal category to eventually calculate the percentage of each of the aforementioned outcomes. The percentages were rounded off to 0 decimal places and were calculated using the formula: 
`=(B2/E2)*100`
Calculating the percentage of eacg outcome gives an overall picture of how the **Plays** campaigns fared overall, whether most were successes or failures or canceled. 

3. Lastly, a line chart was created to show the percentage of the outcomes changing as different fundraising goals were set. The following chart depicts the percentage trendline of the outcomes for different goal levels:

![Outcomes Based on Fundraising Goals](/Resources/Outcomes_vs_Goals.png)


## Summary

- What are two conclusions you can draw about the Outcomes based on Launch Date?
    - Highest number of successful Theater campaigns occur in the month of May across all years
    - The number of successful campaigns start reducing as the year ends, while the number of failed starts to increase during that time

- What can you conclude about the Outcomes based on Goals?
    - The percentage of successful outcomes reduces as the fundraising goal increases, while the reverse is true for failed campaigns (percentage increases as the goals increases)

- What are some limitations of this dataset?
    - The dataset comes without description of certain columns, like Staff Pick and Spotlight, therefore those cannot be used in any form of analysis

- What are some other possible tables and/or graphs that we could create?
    - Tables and graphs could be created to study the outcome trends in different countries for different subcategories (perhaps theatres fare better in some countries over others, whereas movies fare better in some and not in others)
    - Another analysis can be performed on the outcomes based on how many backers each campaign had