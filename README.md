# Stock Analysis Challenge

## Overview
In this challenge, we attempted to refactor the code we wrote throughout Module 2 to loop through all the data one time to improve efficiency and expand the dataset that we could analyze.
This would let us analyze much larger datasets in less time if we wanted to expand the selection.

## Results
The refactored code ended up being much faster than our original, running in a few tenths of a second compared to close to two seconds. While that isn't a huge difference in the grand scheme of things, if we were analyzing 1,200 stocks instead of 12 stocks, it would cut down considerably on wait time.
The code ran so much faster because our For loop looped through all the data at once rather than looping 12 times to grab the corresponding stock each time, as shown below:
![If Then Code]If Then Code.png

The code helped show us that overall, the returns for the list of our stocks were much better in 2017 compared to 2018. However, ENPH and RUN had some great returns in 2018. This might be helpful for our friend to determine how to balance his portfolio find new stocks to invest in. Full year by year results below.
![2017 return]2017 data.png
![2018 return]2018 data.png

## Summary

### Advantages and Disadvantages of Refactoring

Advantages of refactoring are that as we've shown, the code can run much faster and save you time depending on the size of your dataset. Another advantage, especially for a beginner like myself, is that it helps teach you more about the code just by how changes can impact the final product.

Disadvantages might be that you make a mistake and spend more time fixing errors than analyzing anything. If this code is a one time use, it would probably be quicker just to create one that works rather than one that's much more efficient.

### Pros and Cons Application to Refactoring the Original VBA script

In this case, it took my a while to figure out the correct syntax of everything to get the code to loop once and give the same output as our original code. The end result was saving around 2 seconds of wait time.

A pro was that we created a more efficient code and learned a lot about how to write a better VBA script. A con was that our time saved wasn't really worth it for this data set, but it definitely would be if we were dealing with more data.
