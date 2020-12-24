# An Analysis of Various Green Energy Companies
Module 2 of Data Analytics Bootcamp

## Scope
The goal of this project was to take a spreadsheet containing data on various green energy companies' stock performances over the 2017 and 2018 calendar years and, using Microsoft Excel and VBA, take the large amount of data and produce an automated, reproduceable analysis. The analysis was in service of the client, Steve, who sought to help his parents create an investment portfolio based on green energy companies.

## Analysis
The analysis looked at the stock prices of 12 different green energy companies over the years of 2017 and 2018. The scope of the analysis included over 6000 rows of data, and it was deemed prudent to automate the analysis using VBA. At first, the code written to analyze the data used nested for loops to analyze the data for each company, but this resulted in some lines of code being run over 36000 times for a given year. So, using arrays, the code was refactored to iterate through the data once, rather than once per company. 
![Original](resources/Original.png)
![Refactored](resources/Refactored.png)
While the difference in duration is relatively minor on this scale
