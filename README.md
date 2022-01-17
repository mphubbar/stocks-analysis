# Stocks Analysis

## Overview of Project
The purpose of the project was to understand how to create more efficient code using an array instead of having for loops loop continuously through the entire data set.

### Results
Using for loops with arrays proved to be much more efficent than simply using for loops to loop through the entire data set. 

- For example, on the 2017 data set the original code ran in 1.13 seconds as demonstrated by the below graphic. 

![VBA_Slow_Code_2017](https://user-images.githubusercontent.com/93684808/149694517-5423dd04-8560-4efc-b913-4a39cc4805ba.png)

- Conversely, using for loops the code on the 2017 data set ran in 0.42 seconds (see below), which is a significant improvement, especially if we wanted to make the data set even larger.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/93684808/149695098-b5600116-0c7b-4b54-ab2e-05f138883c5c.png)

- The 2018 results were similar as seen in the graphics below with just the for loop first and then using the array with the for loop second. 

![VBA_Slow_Code_2018](https://user-images.githubusercontent.com/93684808/149695445-108bf77b-48b0-4554-b983-9fc6918228a5.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/93684808/149695529-385fef9e-414f-4e39-a044-5f6d507c7f3c.png)

### Summary
- There were several advantages to using refactored code:
     - The outline for the project was already completed saving time
     - It was clear to see the overall intent of the outline and initial code, which made it faster to get started and make progress
     - Some of the more repeatative and less value add code (like formatting) was already complete
- However, there were several disadvantages as well:
     - It took time to read the code and understand the outline before work could be started
     - Some of the naming convention and use of space was not what I would have started with
     - It was easy to gloss over existing code that needed to be updated or alterted slightly, which made bug fixing more cumbersome

- As described above there were significant advantages in terms of efficiency that were gained from using the refactored VBA script that combined arrays with for loops instead of simply using for loops to loop through the entire data set. While the data set in the exercise was relatively small, the difference was still roughly a 50% improvement and it is easy to imagine how the original code would be unlikely to scale as the dataset got larger. 
