# Analyzing Stocks

## Overview of Project: 

## Results: 
Comparison of Stock Performance Between 2017 and 2018

2017            |  2018
:-------------------------:|:-------------------------:
![](https://github.com/vivek-gurumoorthy/stock-analysis/blob/main/Screen%20Shot%202022-07-28%20at%202.59.21%20PM.png)  |  ![](https://github.com/vivek-gurumoorthy/stock-analysis/blob/main/Screen%20Shot%202022-07-28%20at%202.59.58%20PM.png)
- Evidently, the selected stocks as a whole performed significantly worse in 2018 than they did in 2017. This is made clear by the conditional formatting used to highlight positive return values in green and negative return values in red.
- Only two of the twelve selected stocks demonstrated consecutive years with positive return values. These stocks are ENPH (129.5% in 2017, 81.9% in 2018) and RUN (5.0% in 2017, 84.0% in 2018).

Comparison of Previous vs. Refactored Script Performance

Old Script            |  Refactored Script
:-------------------------:|:-------------------------:
![](https://github.com/vivek-gurumoorthy/stock-analysis/blob/main/Old_2018.png)  |  ![](https://github.com/vivek-gurumoorthy/stock-analysis/blob/main/Refactored_2018.png)

## Summary: 
1. What are the advantages or disadvantages of refactoring code?

  a. **Advantages** - Refactoring code makes it more efficient and globally applicable. For example, changing the macro from its previous format that was only meant to use with the 2018 spreadsheet into one that could be used for either 2017 or 2018 made the code more useful. Beyond that, refactoring the code further to make it less reliant on a long and nested for loop made it more run more quickly and delivered analytical results more efficiently.
  
  b. **Disadvantages** - In refactoring, one has to avoid hard-coding certain values in order for this code to be readily applicable to more settings and situations. In doing so, however, writing the code itself can become more abstract and error-prone. Instead of referring to objects of which the values are readily known and viewable, one writing the code must be able to envision how the code is working more theoretically which may be challenging.  

2. How do these pros and cons apply to refactoring the original VBA script?



