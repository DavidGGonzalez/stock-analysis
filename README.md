# VBA Challenge

## Project Objective
After a successful work presented to Steve, he wanted to expand the research for his parent extending the analysis to include the entire stock market for the last few year, forcing me to review the code and come up with any modifications that would improve its performance while keeping results accuracy. An initial proposal was provided to accomplish this task, and it needs to be evaluated to see if it is the best way to write the new code.

## Results
The suggested method resulted to be effective and right for the provided dataset, it was a perfect example of refactoring code looking for better results.

### Suggested Method
As suggested for this task, I've created a subroutine called ***AllStocksAnalysisSuggested***, storing the results in 3 separate arrays for each result, then updating the target worksheet out of the loop and using the `tickerIndex - 1` as the loop top value; the reason for the subtraction is because tickerIndex gets incremented at the end and this would try to read and create and additional row and there is no value for it in the arrays which would return an error.

__2017 - Suggested Method__ (VBA_Challenge_2017.png)

![2017 - Suggested Method](/Resources/VBA_Challenge_2017.png)

__2018 - Suggested Method__ (VBA_Challenge_2018.png)

![2018 - Suggested Method](/Resources/VBA_Challenge_2018.png)

## Summary
So far the results showed by the suggested implementation as compared with the previous class excercise shows a significant improvement, the right code will always be the with the less amount of lines, linear, easy to understand and that does the job better than the previous approach, we would have to test it with a more large data set. Like I've previously said, I would  at least 1 more thing and that is to include the validation of the provided year or worksheet name via the `InputBox` to make sure the worksheet exists.

Thanks a lot to the evaluator for his/her advice which made get back to review my code, that forced me to think further.
