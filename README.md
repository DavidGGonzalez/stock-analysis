# VBA Challenge

## Project Objective
After a successful work presented to Steve, he wanted to expand the research for his parent extending the analysis to include the entire stock market for the last few year, forcing me to review the code and come up with any modifications that would improve its performance while keeping results accuracy. An initial proposal was provided to accomplish this task, and it needs to be evaluated to see if it is the best way to write the new code.

## Results
The suggested method resulted to be more effective and right for the provided dataset, it was a perfect example of refactoring code looking for better results while using less code.

### Suggested Method
As suggested, I've created a subroutine called ***AllStocksAnalysisSuggested***, it stores the results in 3 separate arrays for each result, then updating the target worksheet out of the loop and using the `tickerIndex - 1` as the loop top value; the reason for the subtraction is because tickerIndex gets incremented at the end and this would try to read and create and additional row and there is no value for it in the arrays which would return an error.

__2017 - Suggested Method__ (VBA_Challenge_2017.png)

![2017 - Suggested Method](/Resources/VBA_Challenge_2017.png)

__2018 - Suggested Method__ (VBA_Challenge_2018.png)

![2018 - Suggested Method](/Resources/VBA_Challenge_2018.png)

## Summary
Definetely the refactored code provides a more effective and faster way to get the daily total, starting, and ending values while calculating the return in just 1 pass, avoiding the need to iterate through the entire dataset for every stock we found in it. We might not appreciate a "***big difference***" with the provided dataset but definetely would in a more larger one, perhaps containing the entire stock market.

Refactoring could be time consuming as you may find many challenges that probably are not necessary if the code is running and providing the results, but based in maintainability you will get a more fresh/clear code, easier to read and understand, and of easier to maintain, not to mention performance improvement. Overall, there is at least 1 more thing I would include, and it is the validation of the provided year via the `InputBox` to make sure the worksheet exists, perhaps a way to read every possible stock found in the list as we hard coded them here.
