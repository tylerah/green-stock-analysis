# Stock Analysis Using Excel VBA Macros
Click the following link to view the Excel file: https://github.com/tylerah/green-stock-analysis/blob/main/VBA_challenge.xlsm
## Overview of Project
### Purpose
The purpose of this project was to refactor a VBA script that was originally designed to display data associated with various stocks from green energy businesses. The original code was able to analyze the stock data contained in an Excel worksheet and output both the total daily volume as well as the overall return for the year. The client was so impressed with the worksheet that he wants to be able to use it on the entire stock market. The original code works well for a dozen stocks, but was not written with efficiency in mind. As such, the goal of this project was to optimize the original code so that the script can be used on larger data sets without taking too long to execute.
### Data Displayed by the Excel VBA Script
The VBA script collects information on stocks pertaining to 12 green energy companies. The information collected and displayed includes: the Total Daily Volume and the Return for a given year. The macro allows the user to select which year they wish to retrieve this information for. The displayed data is then formatted with color to easily see which stocks experienced positive and negative returns.

The following image is an example of how the data is displayed after the VBA script runs:

![example_of_display_data](https://user-images.githubusercontent.com/104606662/168720802-8fb872dd-a1de-4b77-8a1a-0dbaebce797f.png)
## Results
### Analysis
The original code ran in roughly .8 to .9 seconds for each year. The refactored code was able to accomplish the same in less than .2 seconds for each year. See below images.

![VBA_challenge_2017](https://user-images.githubusercontent.com/104606662/168720892-f10b98ec-2e6b-49b7-8c57-1081bc0ddbdd.png)

![VBA_challenge_2018](https://user-images.githubusercontent.com/104606662/168720896-648457c1-2e15-4a75-8774-66fc52443637.png)

The original code took almost four times as long because it utilized a nested loop that ran through every row 12 times in order to collect information for each stock (there were 12 in total). The refactored code was able to perform the same function in a single loop by utilizing a set of arrays. See the images below for examples of the code.

The following image displays the original code with the nested loop:
![original_code_nestedloop](https://user-images.githubusercontent.com/104606662/168720951-8ca32433-f8aa-4d60-972f-93c1e962d85c.png)

The following image displays the updated code utilizing a single loop:
![updated_code_singleloop](https://user-images.githubusercontent.com/104606662/168720987-2f0fd948-4f79-4fd7-8b75-6e81c974feca.png)

The following image displays the updated code's array design used to enable a single loop:
![updated_code_arrays](https://user-images.githubusercontent.com/104606662/168721032-cc5e679b-b039-4762-9456-92a52e231d17.png)

## Summary
### Pros and Cons Associated With Refactoring Code
Refactoring is the process of restructuring existing code without fundamentally changing the purpose/function of the code.

The advantages of refactoring code include: increased efficiency of the code (decreasing the time it takes to execute), increased facilitation of debugging, and improved legibility so that it is comprehensible to anyone else who might need to review the code. Improving legibility also allows other developers to see what is going on such that the code can be easily upgraded or adjusted by anyone and not just the original author.

One major disadvtange of refactoring code is the possible disruption of integration with other pieces of code and applications. If one is not careful, it can be easy to unintentionally break integration of one piece of code with another. Additionally, refactoring code takes time and money. However, as the function of the code is not changed the end user may not appreciate the benefits of refactoring. Because of this, one must always evaluate whether the time and money associated with refactoring is creating enough of a benefit to justify the cost.
### Advantages of Refactoring the Stock Analysis
The biggest benefit from refactoring the VBA script utilized in this stock analysis is the increase in efficiency. As mentioned above, the original code took almost 0.8 seconds to run while the refactored code took just less than 0.2 seconds. While this is a negligible difference on the small data sample used in this project, it would be noticeable on larger datasets. As the client wants to use this script on datasets containing the entirety of the stock market, it is possible that the optimized code could save significant time for the end user.
