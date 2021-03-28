# UCB-Excel-Project
## Green Stock Investment with VBA

## Overview of Project
I used Visual Basic for Application  (VBA) to analyze a stocks data with a goal to choose a profitable stock for investment. I used prescarpped dataset of 12 stocks and analyzed their performance (Total Volume and Return) over two years (2017 and 2018). The client is especially interested in knowing the performance of DQ stock. 

The results are obtained by refactoring a code and the results were compared with the original code. 

## Results
### 1. Stock performance:
As we can see from the pictures below:
- Year 2017 was  over all a good year as compared to 2018. 
- DQ was the best performing stock in 2017 but it didn't do so well in 2018. 
- ENPH has consistenty performed well over two years. 

<img width="304" alt="Results 2017" src="https://user-images.githubusercontent.com/69255270/112765896-345ffd00-8fc4-11eb-87e2-9b59a3dc8ebe.png">                 <img width="304" alt="Results 2018" src="https://user-images.githubusercontent.com/69255270/112765901-35912a00-8fc4-11eb-909c-230e30fe8c53.png">

In nutshell, investing in ENPH instead of DQ will be a good decision. 

### 2. Refactoring Performance:

The code ran much faster when refactored (images below).

#### 2(a) Run time with Original Code:
<img width="234" alt="Original Code Timing" src="https://user-images.githubusercontent.com/69255270/112768216-ca4d5500-8fcf-11eb-84cb-a384b87a34cb.png">          <img width="247" alt="Original Code Timing1" src="https://user-images.githubusercontent.com/69255270/112768219-cde0dc00-8fcf-11eb-8a3c-2d65c6ba3e27.png">

#### 2(b) Run time with Refactored Code:
<img width="304" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/69255270/112768550-60ce4600-8fd1-11eb-9356-6c5536253c19.png">           <img width="304" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/69255270/112768477-18168d00-8fd1-11eb-9966-a129058c3e78.png"> 

This happens because in refactor code, I used arrays that eliminate the need of nested loops that we used in the original code. As you can see in the images below that refactored code is easier to understand and execute, hence it takes less time. 

##### 2(c) Original Code
<img width="490" alt="original Code" src="https://user-images.githubusercontent.com/69255270/112768054-e7cdef00-8fce-11eb-9a77-64d7ee474d96.png">

###### 2(d) Refactored Code
<img width="500" alt="Refactored Code" src="https://user-images.githubusercontent.com/69255270/112768419-e998b200-8fd0-11eb-8371-3e5c26929c95.png">

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

