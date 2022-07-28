# Analyzing Stocks

## Overview of Project: 
### Purpose

### Background


## Results: 
Comparison of Stock Performance Between 2017 and 2018

2017            |  2018
:-------------------------:|:-------------------------:
![](https://github.com/vivek-gurumoorthy/stock-analysis/blob/main/Screen%20Shot%202022-07-28%20at%202.59.21%20PM.png)  |  ![](https://github.com/vivek-gurumoorthy/stock-analysis/blob/main/Screen%20Shot%202022-07-28%20at%202.59.58%20PM.png)
- Evidently, the selected stocks as a whole performed significantly worse in 2018 than they did in 2017. This is made clear by the conditional formatting used to highlight positive return values in green and negative return values in red.
- In 2017, the stock with the highest total daily volume was SPWR with 782,187,000 and the stock with the lowest total daily volume was DQ with 35,796,200. The stock with the highest return value was DQ with 199.4% and the stock with the lowest return value was TERP with -7.2%. Eleven out of the twelve selected stocks had positive return values. 
- In 2018, the stock with the highest total daily volume was ENPH with 607,473,500 and the stock with the lowest total daily volume was AY with 83,079,900. The stock with the highest return value was RUN with 84.0% and the stock with the lowest return value was DQ with -62.6%. Two out of the twelve selected stocks had positive return values.
- Only two of the twelve selected stocks demonstrated consecutive years with positive return values. These stocks are ENPH (129.5% in 2017, 81.9% in 2018) and RUN (5.5% in 2017, 84.0% in 2018).

Comparison of Previous vs. Refactored Script Performance

Old Script            |  Refactored Script
:-------------------------:|:-------------------------:
![](https://github.com/vivek-gurumoorthy/stock-analysis/blob/main/Old_2018.png)  |  ![](https://github.com/vivek-gurumoorthy/stock-analysis/blob/main/Refactored_2018.png)
- The refactored script yields the same output as the original script, but reaches this output much more efficiently. This is shown by the difference in the time displayed in the Excel message box between the previous version and the new and refactored one. While the older version of the script takes nearly an entire second to reach the output, the refactored version reaches this same result more than four times faster. Code that can reach a desired end result at a faster rate is much more desirable and easily usable than code that takes significant time to perform a task, and for this reason the refactored code is preferable.

## Summary: 
1. What are the advantages or disadvantages of refactoring code?

  a. **Advantages** - Refactoring code makes it more globally applicable. When one writes a VBA script, while it is useful for that script to successfully solve one specific task, it would bring more value to write that script in a way in which it can also be applied to other similar tasks for future use. Additionally, refactoring makes the code itself reach its end goal more quickly, thereby allowing needed insights to be reached by less time-costly means.
  
  b. **Disadvantages** - In refactoring, one has to avoid hard-coding certain values in order for this code to be readily applicable to more settings and situations. In doing so, however, writing the code itself can become more abstract and error-prone. Instead of referring to objects of which the values are readily known and viewable, one writing the code must be able to envision how the code is working more theoretically which may be challenging.  

2. How do these pros and cons apply to refactoring the original VBA script?

  a. **Pros** - Changing the macro from its previous format that was only meant to use with the 2018 spreadsheet into one that could be used for either 2017 or 2018 made the code more useful and applicable to other potential years of analysis. Beyond that, refactoring the code further to make it less reliant on a long and nested for loop made it more run more quickly and delivered analytical results more efficiently.
  
  b. **Cons** - While the script itself is more useful and runs more quickly, refactoring the script was somewhat challenging because it required an understanding of how to index into and refer to items in an array. Rather than referring to a ticker, volume, starting price, or ending price object directly, accessing these individual objects required indexing into the arraay in which these objects were located. While at first this proved to be a bit difficult, visualizing the data structures enabled me to realize how the script should be properly constructed. 


