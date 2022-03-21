# Wolf of Wall Street with Excel VBA

## Project Overview
My friend Steve asked me to help him analyze stock data for the stock market.  I already assisted Steve with a Microsoft VBA code that sorted through the data of twelve green energy stocks.  Steve had two years worth of stock data; for years 2017 and 2018.

### Purpose
The purpose of this project is to refactor the VBA code to make it faster and more efficient.  Again, I was tasked with sorting through two years of stock market data, 2017 and 2018.  The goal for the VBA code was to retrieve the Tickers, the Total Daily Volume, and the Return on each stock.  The Return section was color coded to make it easy to show if the return was profitable or not.  

## Analysis
To start this project, I copied and pasted the below original code, with sudo-code, in the VBA editor.
#### Original Code
<img width="639" alt="Original_Code" src="https://user-images.githubusercontent.com/99417460/159197724-dab2da1f-ee6f-42c6-98ed-39ba9b1dae87.png">
I then refactored the code, pictured below, to be more readable and more efficienct.
#### Refactored Code
<img width="585" alt="Refactored_Code_Pt_1" src="https://user-images.githubusercontent.com/99417460/159198032-4ed046d7-c974-4798-a7d0-f7bf3e22705c.png">
<img width="587" alt="Refactored_Code_Pt_2" src="https://user-images.githubusercontent.com/99417460/159198045-aa3253eb-75ab-4060-bc79-46e17eb9f328.png">

### Results
#### Original Timer
<img width="256" alt="Pre_Refactor_2017" src="https://user-images.githubusercontent.com/99417460/159198193-b61de232-3e03-4970-842d-6b853346dcb8.png">
<img width="260" alt="Pre_Refactor_2018" src="https://user-images.githubusercontent.com/99417460/159198203-85465771-ad44-45b7-90f3-b07a97691b4d.png">

#### Refactored Timer
<img width="259" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/99417460/159198224-e907d787-f459-47b7-8198-e5a031ed20c1.png">
<img width="259" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/99417460/159198246-52c677ce-d124-4186-8e6c-6e188c6889b6.png">

As one can see, the refactored code analyzed the data much more quicker and efficient than the origianl code, but about two-tenths of a second.  That may not sound like much, but with a larger amount of data to analyze, the refactored code can save much more time than the original.

## Summary
### Advantages and Disadvantes of Refactoring Code in General
Refactoring code can make code look more clean and organized, therefore making it easier to read and sort through.  Another advantage are that it can help the code to run faster.  Refactoring can also fix bugs.  On the other hand, refactoring can introduce bugs to our code.  Refactoring also takes time, which we may not have.

### Advantages and Disadvantages of Refactoring This VBA Script
The obvious advantages of refactoring this VBA script are that it was faster and more efficient.  The code was easier to read and follow through.  Therefore the end product was that the code processed around two-tenths of a second faster than the original code.  The disadvantage was the time it took to refactor the code to where Excel didn't give an error message.
