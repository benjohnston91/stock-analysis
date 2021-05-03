# Stock Analysis Using VBA in Excel

### Overview of Project

- The purpose of this project is to utilize the skills we've developed in Excel using VBA to refactor existing code in order to make it more efficient. This project displays the time needed for our macro to analyze numerous stocks' annual volume and performance, showing our refactored code produced the same results in less time.

### Results

- Given that our macro presents a message box with the time elapsed, it is very clear to see how our refactored code increased the efficiency. Using our original code that we developed throughout Module 2, it took over half a second to run our macro for both 2017 and 2018:

<img width="1170" alt="2017_Original_Code" src="https://user-images.githubusercontent.com/82119497/116933522-a8c83480-ac31-11eb-9680-1bdd464772a3.png">

<img width="1181" alt="2018_Original_Code" src="https://user-images.githubusercontent.com/82119497/116933538-acf45200-ac31-11eb-927a-8216f814bf0b.png">

Now, let's compare that to our refactored code:

<img width="1191" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/82119497/116933668-ddd48700-ac31-11eb-9708-8270ab5b0c47.png">

<img width="1407" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/82119497/116933679-e0cf7780-ac31-11eb-9cfe-4f52cbbfd40c.png">

We managed to reduce the time needed to less than two tenths of a second! 

### Summary

- Refactoring code has both advantages and disadvantages. The advantage is that you are making your existing code more efficient. As demonstrated in the results section above, our refactored code produced the same results as the original code in less time. If this is code you will be using regularly, that time adds up and will pay off in the long run. The disadvanatge is that the original code did produce the same results, just in a less efficient manner. So if this analysis is a one-time event, it may be tougher to recognize the benefit of the refactored code.

- These pros and cons of refactoring apply to the original VBA script because of the increase in efficiency demonstrated in the results section. The new refactored code runs much more quickly than the original VBA script, so if this analysis will be run repeatedly, time will be saved.
