# VBA of WallStreet

## Overview of Project

Our objective is to refactor the code provided and create a more efficient and simple to use macro which will run the analysis for us in a shorter period of time with no errors in the code while still returning correct analysis of the trends in the stocks listed. In my code, I added the option for users to clear the spreadsheet using a macro as well in order to keep sheet clean and easy to use.

### Purpose

The purpose of this project is to use VBA code to loop through the Green Stock data provided in order to analyze the stock market during the years 2017 and 2018 and determine which stocks are likely to be the best investments based on previous years data.

## Results and Analysis

Based on the images of the stock analysis that I have included in the resources folder, we can see that in 2017 all stocks except for TERP produced positive returns. In 2018, only two stocks returned positive results, ENPH and RUN. We can also see that many of the stocks in 2017 that were successful in their return in value were traded in a significantly higher daily volume, for example AY had a daily volume of 136,070,900 in 2017 versus its daily volume of 83,079,900 in 2018. Based on this information, we can determine that based on the decline in daily volume and returns from 2017 to 2018, this is not the best set of stocks to invest in at the moment. 

I have included photos in the resources folder that display the execution time of the original code and refactored code for both 2017 and 2018. Based on the data shown, we can conclude that the refactoring of the code improved the execution time. 

## Results

- What are the advantages or disadvantages of refactoring code?

Refactoring code is advantageous for both the coder and the user. By refactoring code we can reduce the time it takes for macros to execute their functions which can be especially helpful when analyzing extremely large datasets that would otherwise take significant amounts of time to go through. Refactoring code also allows us the opportunity to clean up and organize code differently, making it easier for users to understand and potential refactor themselves again in the future.

There are several potential disadvantages of refactoring code, the first of which is the time spent refactoring, as the process can take significantly longer for users with less experience or when working with messy code. The more there is to refactor, the more complicated the task becomes and you run the risk of making an already working macro stop working. I found several times that I would change one thing, everything would stop working, and I would have to load a previous saved version of my work. 


- How do these pros and cons apply to refactoring the original VBA script?

These pros and cons are clear while I was working on refactoring the original VBA script. By refactoring the code, I was able to gather a better understanding of why it is important to organize our code in specific ways with headings for ease of access. There were many times that I felt lost or needed to find a specific line and was able to use the organization added to help me identify the problem with my code and make changes. I was also made very aware of the cons of refactoring, as it took me a significant amount of time to feel comfortablewith the process of refactoring the code and I made many, many mistakes along the way. I experienced several issues with changing one letter somewhere and finding that my whole code was broken and I would have to open a previous version. 

