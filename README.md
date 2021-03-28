# UCB-Excel-Project
## Green Stock Investment with VBA

## Overview of Project
I used Visual Basic for Application  (VBA) to analyze a stocks data with a goal to choose a profitable stock for investment. I used prescarpped dataset of 12 stocks and analyzied their performance (Total Volume and Return) over two years (2017 1nd 2018). The client is especially interested to know if DQ is a good staock to invest in.

The results are obtained by refactoring a code and the results were compared with the original code. 

## Results
### 1. Stock performance:
As we can see from the pictures below:
- Year 2017 was a good year over all as compared to 2018. 
- DQ was the best performing stock in 2017 but it didn't do so well in 2018. 
- ENPH has consistenty performed well over two years. 

<img width="304" alt="Results 2017" src="https://user-images.githubusercontent.com/69255270/112765896-345ffd00-8fc4-11eb-87e2-9b59a3dc8ebe.png">

<img width="304" alt="Results 2018" src="https://user-images.githubusercontent.com/69255270/112765901-35912a00-8fc4-11eb-909c-230e30fe8c53.png">

In nutshell, investing in ENPH instead of DQ will b a good decision. 

### Refactoring Performance:

The code ran much faster when refactored (images below).

<img width="304" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/69255270/112767924-38911800-8fce-11eb-9adc-202011edd78f.png">

<img width="304" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/69255270/112767930-3b8c0880-8fce-11eb-895c-fb93ae0bba2a.png">

This happens because in refactor code, I used arrays that eliminate the need of nested loops that we used in the original code. As you can see in the images below that refactored code is much easy to understand and execute, hence it takes less time. 

<img width="490" alt="original Code" src="https://user-images.githubusercontent.com/69255270/112768054-e7cdef00-8fce-11eb-9a77-64d7ee474d96.png">

<img width="500" alt="Refactored Code" src="https://user-images.githubusercontent.com/69255270/112768055-eac8df80-8fce-11eb-9a2b-2f0ce8452aab.png">

## Summary: 
In this project the results are obtained by refactoing the code and then comparing the results with the original code. 
### Advantages of refactoring code: 
- It improves the design of the code.
- It makes the code run faster
- it helps debugging easy
- it improves the understanding of the code/programming.

### Disdvantages of refactoring code:
- It can be time consuming
- It doesn't provide additional insight into data
- It doesn't improve the features of the output. 
- it is not necessary if the program is already running efficiently. 
-
### Pros and cons apply to refactoring the original VBA script?

