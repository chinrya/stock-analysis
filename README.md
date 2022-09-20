# stock-analysis-challenge

## Overview of Project:
  
  The purpose of this analysis was to expand our knowlede on VBA, and put into practice more advanced techniques, such as nested for loops, use of several different variable types, and coloring text. We edited a code to be able to pore through thousands of rows of data, to pull volumes and starting and ending prices of respective tickers, for whichever year the user requested. 
  
## Results:

  After running both the original code and the refactored code, I was greatly surprised to see the refactored code ran about four times as fast as the original code, even though there were more lines of code. 

<img width="307" alt="Original" src="https://user-images.githubusercontent.com/109634784/191174093-bc4982fe-3a59-46a0-947d-d431fa865a21.png"> <- Original Code

<img width="306" alt="Refactored" src="https://user-images.githubusercontent.com/109634784/191174130-6186f95a-c655-460c-a3b7-5d5d6a3018ee.png"> <- Refactored Code

I would have thought that because the refactored code I wrote was written sloppily, as I had to troubleshoot at several points, it would run much slower, but it instead seemed much more streamlined. For example, after a lot of Googling, I decided to use ReDim,as shown below, in an attempt to debug an issue, which worked, but it seemed not very efficient in terms of memory, so I would have thought it would slow the process, making it longer than the original. 
  
    Dim tickerVolumes() As Long
    
    ReDim tickerVolumes(RowCount)
    
    Dim tickerStartingPrices() As Single
    
    ReDim tickerStartingPrices(RowCount)
    
    Dim tickerEndingPrices() As Single
    
    ReDim tickerEndingPrices(RowCount)


## Summary:

  Refactoring code in general seems to have noticable advantages and disadvantages to me, a big disadvantage for me personally is not knowing exactly what the original coder was thinking at the time the wrote the code. This could potentially affect variables you wanted, which is an easy fix, but still could cause issues with larger codes. A strong advantage, however, is that you have the groundwork laid out, and if you can easily understand the code, this can expidite the process, as all you would need to do is update the code with whatever changes you need.
  
  For this VBA challenge specifically, a strong advantage is that some of the parts I was slightly confused about were already included, which saved me a decent bit of headache trying to remember the exact syntax and whatnot. This also provided a framework in which I could constantly reference to remember how to do certain functions. The one disadvantage I personally experienced was that I accidentally used variable that slightly differed from what as already written, as in my head I would have writting them with another name, which caused quite a bit of confusion for a moment, before I figured it out. 
