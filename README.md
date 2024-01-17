# VBA-challenge
bootcamp module 2

Some notes based on feedback on original grading
-I'm not sure what the "tickerIndex" variable is supposed to be. If there is something to count the total number of stocks evaluated, I am using a variable named "StockNum" for this.
-I am not sure what the factored and refactored code is referring to. Unless I missed something, the challenge had me create my own code to execute the task required.
-There was not a year 2017 in the provided data set so I am not able to make that comparison between 2017 and 2018. Some of the other specific feedback regarding is still unclear to me as there are not specific analysis instructions on the challenge page or in the source files
-I am now adding the files that I tested with and the multiple year stock data resutlts that I got this for this submission

Code description

    Outside code used:
    This code uses the for loop strategies used in 2.3's activity 6 checking if we are looking at a row with a new name in column 1 (credit card name vs stock ticker) and 7 for looping through each worksheet in a workbook.

    The program will repeat perform a loop for all pages in the worksheet. The following will be performed across all pages:

        Initialization:
            Variables are initialized to track the following values:
            -Current stock being tracked in the main loop
            -Total volume of the current stock bieng tracked
            -Starting value of the year for the stock
            -Final value of the year for the stock
            -Yearly change of the stock
            -Percentage change of the stock
            -total number of rows in the sheet
            -number of stocks checked in the sheet
            -stocks with the greatest % gain, % loss, and volume traded as well as the names of these stocks

            Labels for our final results are also inserted into the page

        Main loop:

            The main loop of the macro will first use an if statement to determine whether or not the next row is a different from the last stock being counted for. As the source data file is already sorted alphabetically by ticker and then by date within the year. Because of this, we can will only need to check the current row against the current stock being tracked in the loop and then treat the first row as the starting value of the year if we are starting to track a new stock.

            If the current row is for the same ticker as the loop previously checked, the loop will simply set the current row as the year end value and the volume total will be updated.

            If the current row is for a new stock ticker, the yearly end value can be retrieved. This can be used to calculate the yearly change and percentage change. As long as it isn't the first stock being checked, the ticker, yearly change, percentage, and total volume of traded stocks. The yearly change is checked if it is positive and negative and that cell will be color coded accordingly. The cell containing the percentage is then formatted as a percentage. The current stock's yearly change and percentage change are compared with the greatest gain, greatest loss, and greatest traded volume. If it surpasses any of these values, it will replace it. Finally, the number of unique tickers checked this year will be incremented. Then the new ticker will be updated and neccesary variables are reset for the new stock.

            Once the loop reaches the end fo the page, the final results will be written to the page from variables. The values are
            -Name of the stock with greatest gain
            -Name of the stock with the greatest loss
            -Name of the stock with the most traded volume
            -Greatest gain achieved which will be formatted as percentage
            -Greatest loss achieved which will be formatted as percentage
            -Greatest volume of stock traded

Results

    Regarding the time of execution, the marco runs almost instantaneously with the alphabetical testing file. When running it on the multiple year stock data page, it takes around 30 seconds.

    Some observations on the results:
    -RKS saw the greatest percent loss for both 2018 (-90%) and 2019 (-91%). While not the greatest loss of 2020 it still saw a percentage loss of 88%