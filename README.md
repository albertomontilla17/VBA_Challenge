# VBA_Challenge
## OVERVIEW OF PROJECT

### In this project, I'm working with a given dataset of stocks with the purpose of creating an automated extraction of data for a better understanding of the market.

## ANALYSIS AND CHALLENGES

### The challenges for this project:
    -Prepare the dataset
    -Refactor the code that we used before 
    -Create an easy enviroment for the users
    -Create the resources for the repository

Steve needs our help to invest in the stock market and thanks to VBA we are able to create this automatic process

## RESULTS

In the All stocks Analysis we can see 2 buttons where we can run the analysis for the year 2017 or 2018

![VBA_Challenge_2017](https://github.com/albertomontilla17/VBA_Challenge/blob/main/VBA_Challenge_2017.png)
![VBA_CHallenge_2018](https://github.com/albertomontilla17/VBA_Challenge/blob/main/VBA_Challenge_2018.png)
## SUMMARY

**1. What are the advantages or disadvantages of refactoring code?**

You need to perform code refactoring in small steps. Make tiny changes in your program, each of the small changes makes your code slightly better and leaves the application in a working state.

**Disadvantages:**

> - A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
> - A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions.
> - A complex unstructured code is usually best to split in several functions. 
> - Refactoring process can affect the testing outcomes. 


**Advantages:**
> - Logical errors easily appear in well structure code that contains nested conditionals and loops. 
> - In our case, using Excel flow displays program logic in a more comprehensible manner, not tied to the order that the underlying code is written.
> - VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.

**2. How do these pros and cons apply to refactoring the original VBA script?**

> Improving or updating the code without changing the software’s functionality or external behavior of the application is known as code refactoring.
Now, let's think about something, **What happens after a couple of days or months yo need to troubleshoot your code? Is it complicated? Is it hard to understand?** If yes then definitely you didn’t pay attention to improve your code or to restructure your code. 

***We need to consider the code refactoring process as cleaning up the orderly house.*** 
*Unnecessary clutter in a home can create a chaotic and stressful environment.* - The same goes for written code. 

A clean and well-organized code is always easy to change, easy to understand, and easy to maintain. You can avoid facing difficulty later if you pay attention to the code refactoring process earlier.



