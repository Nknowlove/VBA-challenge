# Assignment task
Create a script that contains tickers, quarterly changes, percentage changes, and total stock volume for each quarter. 

Loop through for the worksheet data and find the greatest increase, decrease, and total stock volume, and show tickers respectively.

Use conditional formatting to highlight positive change in green and negative change in red.

Make sure that VBA script is able to run on each quarter's data.


## Instructions

1. Use the range fomular to creat ticker, quarterly change, percentage change and so on.

2. To remove the duplicate in ticker. In order to use VBA way, I learn it from websit, how to use dictionary to do this. Because dictionary contains two arrays, which are key and item. The key is unique, and only one value. When I figured out how to build the dictionary, I use if else syntax to make column I. If ticker in dictionay not exist, then add into it. Therefor all ticker value in column is unique.

3. At the beginning, I use I("i,i") to define the range for loop. When running the macro, the value was not correct. Hence, I search on ChatGPT how to write as range from i2 to last cell(syntax is different from the class, because I did this before the 3rd class).

4. For quarterly change and percentage change, I create two dictionaries. I use if else to define that ticker is key and price is item. If ticker does not exist, then add it into dictionary and record the openprice in that row as the beginning open price. One the other hand, as looping trhough from i2 to lastrow, close keep update till to the very last one as final close price. Then I get the quarterly change. The percentage is the value that quarterly change over the beginning open price. There is a liitle bit tricky when I first time deal with the qurately change data. Due to some date is not normal data format. Therefore, when I first time run macro, the value did not come out and only got an error. I searched how to slove it. Finally I got an formula how to fix it. Hence, the date modification is not my original, I just borrowed and reorgazied. After that, I use conditional and interior.colorindex to make the green and red.

5. For the total stock volume, create the dictionary again. Loop through the key(ticker). If does not exist then add it into dictionary, and record volume as itme of the dictionary.Otherwise, sum all looped row values.  

6. For the greatest increase, decrease and volume, I use the same method. I loop through percentage change and total stock volume respectively. Using the maxvalue and minvalue to get both values. Then, I use function called offset, which could shift cell based on current cell. Finally I got the ticker and finish all the assignment. There was a little struggling when I used total stock volume to get sticker. First time I set maxvalue equal to cell.value, I got error" out of merroy". I searched long time and according to one suggestion from websit and rewrite that part again. Whee! Problem solved.

## Declaration
During the task, I use some syntax, fomula and functions. Some of them from class, some of them from web search. When I got an error or had no idea what is going on, I chat with chatGPT and got some suggestions or debug for codings. I figured the problem out first and understood it, and then typed code down. From this view, my work might be have some similarity part with others. I am a new learner on coding, have to do lots of research and reading for problme sloving.
