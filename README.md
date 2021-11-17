# VBA Challenge

## Project Objective
After a successful work presented to Steve, he wanted to expand the research for his parent extending the analysis to include the entire stock market for the last few year, forcing me to review the code and come up with any modifications that would improve its performance while keeping results accuracy. An initial proposal was provided to accomplish this task, and it needs to be evaluated to see if it is the best way to write the new code.

## Results
I went ahead and tried the suggested method since I needed to make sure any modification done (__Code Refactoring__) would produde better performance improvements while running the code.

### Suggested Method
As suggested for this task, I've created a subroutine called ***AllStocksAnalysisSuggested***, basically following all steps, it does the job providing us with the results, and in a better time as compared with our initial project (see screenshots below). Running time average between the 2 worksheets was __0.8__ seconds.

__2017 - Suggested Method__ (VBA_Challenge_2017_SuggestedMethod.png)

![2017 - Suggested Method](/Resources/VBA_Challenge_2017_SuggestedMethod.png)

__2018 - Suggested Method__ (VBA_Challenge_2018_SuggestedMethod.png)

![2018 - Suggested Method](/Resources/VBA_Challenge_2018_SuggestedMethod.png)


### Code __Refactored__
After doing some research on how to improve execution time of a VBA code, another subroutine was created holding the refactored code called ***AllStocksAnalysisRefactored***, I've found that one of the biggest factors to decrease code proformance was the interaction between the code and the worksheet, so I've decided to use a [multidimensional array of type variant](https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-arrays) to hold the entire worksheet and loop through it and just activating the target worksheet (***All Stocks Analysis***) to update and format final values. Running time average between the worksheets was __0.5__ seconds.


__2017 - Refactored__ (VBA_Challenge_2017_Refactored.png)

![2017 - Code Refactored](/Resources/VBA_Challenge_2017_Refactored.png)

__2018 - Refactored__ (VBA_Challenge_2018_Refactored.png)

![2018 - Refactored](/Resources/VBA_Challenge_2018_Refactored.png)


## Summary
It is hard to say that a difference of __0.03__ seconds between the 2 code approaches would impact the overall performance with the provided data set, but definetely would if the data set will include the entire stock market as suggested by the initial request; to be clear with the advantage and disadvantage of each sub routines, we would have to analyze other factors like memory utilization, code re-usability, functions, etc. Overall, there is at least 1 more thing I would include, and it is the validation of the provided year via the `InputBox` to make sure the worksheet exists.