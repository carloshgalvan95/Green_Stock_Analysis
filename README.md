# Securing a green future in Stocks

## Overview of the project
---
Steve’s parents are passionate about green energies, they believe that, as fossil fuels get used up, there will be more and more reliance on alternative energy production. However, Steve’s parents haven’t done much research and have instead decided to invest all their money into DAQO New Energy Corporation a company that makes silicon wafers for solar panels. Steve is concerned about diversifying their funds. He wants to analyze a handful of green energy stocks in addition to DAQO stock.

The main purpose of this analysis is to have more information and a point of comparison to base the decision of which company or companies to invest in on more than just a guess. We will be comparing eleven green energy companies based on three main indicators.

### Total Daily Volume of Trades

To ensure that the price that we are going to be evaluating reflects as accurately as possible the actual value of the stock we will be taking into consideration this first data point. Just as the central limit theorem states,

>*In probability theory, the* **central limit theorem (CLT)** *states that the distribution of a sample variable approximates a normal distribution (i.e., a “bell curve”) as the sample size becomes larger, assuming that all samples are identical in size, and regardless of the population's actual distribution shape.*
*-* *Ganti, A. (2022, July 8). What is the Central Limit Theorem (CLT)? Investopedia.*

The same idea applies to our stock value if the total daily volume of trades increases so as the accuracy of the value projections.

### Stocks in, stocks out

Our two other our two other data points are quite simple to define, the **stock value starting price** and the **ending price**, with this three data points we can get a general idea of how every company performed throughout the years that we will be using as our points of reference. For now, those years are going to be **2017** and **2018**.

### VBA as a powerful stock analysis tool

For this project we will be using Visual Basic for applications in Excel, to automate the process and ensure the analysis can be done in the future, as our predictions will vary quite a lot in a field as volatile as the stock market. We will program a macro that allows us to run over all the rows of information that we have to obtain and compare in a final table each of the eleven prospect companies to invest in.

Let’s first look at a normal row in our dataset to delimitate the information we will need.

| Ticker | Date       | Open  | High  | Low   | Close | Adj Close | Volume |
|--------|------------|-------|-------|-------|-------|-----------|--------|
| AY     | 2017-01-03 | 19.49 | 19.64 | 19.24 | 19.47 | 16.80219  | 309500 |

This is the first row of our 2017 dataset, and just to sum things up, we are going to need to be able to extract the information for three columns. **Ticker (A,2)**, **Close(F,2)**, **Volume(H,2)**.

**Ticker(A,2)** is going to allow us to identify which company this specific row of information is coming from, we have a Ticker for every company we will be analyzing.

**Close(F,2)** we will be using this to define two of our data points **stock value starting price** and **ending price**. We will need to store these two pieces of information for every single ticker to then obtain the return percentage.

**Volume(H,2)** finally the volume column gives us the volume of transactions for every company in every row of the dataset, amount that we will need to be able to sum and store.

### Turning words into actions
---
Now, let’s brainstorm, what do we need to achieve all of that? Let’s break it down into questions:

#### What data will we be analyzing?

Simple enough, We will be working with the spreadsheets for both 2018, and 2017, for every Ticker (company) and to be able to store those 3 data points in a table for visualization.

That means we need:

1.  The year to run the analysis macro in
2.  The tickers (company) that we will be analyzing
3.  Create a table to output the final 3 data points established

And two other variables, we want to also be able to evaluate how efficient our macro code is, so let’s define how long it takes to run it:

4.  Define a variable to store the start time
5.  Define a variable to store the end time

```

Sub AllStocksAnalysisRefactored()

Dim startTime As Single
Dim endTime As Single

yearValue = InputBox("What year would you like to run the analysis on?")

startTime = Timer

'Format the output sheet on All Stocks Analysis worksheet

Worksheets("All Stocks Analysis").Activate
Range("A1").Value = "All Stocks (" + yearValue + ")"

'Create a header row

Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"

'Initialize array of all tickers

Dim tickers(12) As String
tickers(0) = "AY"
tickers(1) = "CSIQ"
tickers(2) = "DQ"
tickers(3) = "ENPH"
tickers(4) = "FSLR"
tickers(5) = "HASI"
tickers(6) = "JKS"
tickers(7) = "RUN"
tickers(8) = "SEDG"
tickers(9) = "SPWR"
tickers(10) = "TERP"
tickers(11) = "VSLR"

```

#### What do we need the macro to achieve?

We need to be able to retrieve 3 data points for every ticker, we previously defined what those 3 are, so we are going to a use for loops to verify, for every row of the dataset which value corresponds to what ticket and be able to store and sum that volume into a defined variable, we will use the array of all the tickers to loop through all the tickers for every row.

Then, to get the starting and ending price we will verify if the row we are currently evaluated is the first or the last row of information for every ticker and store that data.

```

'Get the number of rows to loop over

RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'1a) Create a ticker Index

Dim tickerIndex As Single

tickerIndex = 0

'1b) Create three output arrays

Dim tickerVolumes(12) As Long
Dim tickerStartingPrice(12) As Single
Dim tickerEndingPrice(12) As Single

''2a) Create a for loop to initialize the tickerVolumes to zero.

For i = 0 To 11

  tickerVolumes(i) = 0

Next i

''2b) Loop over all the rows in the spreadsheet.

For i = 2 To RowCount

    '3a) Increase volume for current ticker

    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

    '3b) Check if the current row is the first row with the selected tickerIndex.

    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i, 1).Value \<\> Cells(i - 1, 1).Value Then

        tickerStartingPrice(tickerIndex) = Cells(i, 6).Value

    End If

    '3c) check if the current row is the last row with the selected ticker

    'If the next row's ticker doesnt match, increase the tickerIndex.

    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i, 1).Value \<\> Cells(i + 1, 1).Value Then

        tickerEndingPrice(tickerIndex) = Cells(i, 6).Value

        '3d Increase the tickerIndex.

        tickerIndex = tickerIndex + 1

    End If

Next i

```

Already having those 3 data points, the only thing left is just doing math

```

'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.

For i = 0 To 11

    Worksheets("All Stocks Analysis").Activate

    Cells(i + 4, 1).Value = tickers(i)

    Cells(i + 4, 2).Value = tickerVolumes(i)

    Cells(i + 4, 3).Value = tickerEndingPrice(i) / tickerStartingPrice(i) - 1

Next i

```

Formating the output table

```

'Formatting

Worksheets("All Stocks Analysis").Activate

Range("A3:C3").Font.FontStyle = "Bold"
Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("B4:B15").NumberFormat = "\#,\#\#0"
Range("C4:C15").NumberFormat = "0.0%"
Columns("B").AutoFit

```

And for better visualization, applying conditional formatting through the macro

```

dataRowStart = 4
dataRowEnd = 15

For i = dataRowStart To dataRowEnd

    If Cells(i, 3) \> 0 Then

        Cells(i, 3).Interior.Color = vbGreen

    Else

        Cells(i, 3).Interior.Color = vbRed

    End If

Next i

```

We end up the macro by getting our end timer run time and inputting in a message box to be able to actually see how it performed in terms of code efficiency.

```

endTime = Timer
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

```
### Analyzing our results

After running our analysis for 2017 and 2018, we get this results:
2017                       |  2018
:-------------------------:|:-------------------------:
![2017_Table](https://github.com/carloshgalvan95/stock-analysis/blob/main/Resources/2017_table.png)  |  ![2018_Table](https://github.com/carloshgalvan95/stock-analysis/blob/main/Resources/2018_table.png)

First let’s evaluate our first column here, **Total Daily Volume**, this will tell us, how accurate the return we are getting actually is for future years. An easier way to visualize this, given that we don’t really have a set amount minimum of total daily volume to consider the return within our desired ranges of accuracy, would be using a treemap.

2017                       |  2018
:-------------------------:|:-------------------------:
![DailyVolume2017_Treemap](https://github.com/carloshgalvan95/stock-analysis/blob/main/Resources/DailyVolume2017_Treemap.png)  |  ![DailyVolume2018_Treemap](https://github.com/carloshgalvan95/stock-analysis/blob/main/Resources/DailyVolume2018_Treemap.png)

#### Finding order within chaos

Based on this information, we see that, for 2017 We could trust that the accuracy of the return could be better for subsequent years for SPWR and FSLR tickers and ENPH, SPWR, RUN and FSLR for 2018.

Following our desired accuracy parameters, this means that we should be mainly focusing on this tickers as our possible investment options. Stock prices fluctuations vary on thousands of correlated variables that makes impossible to consider all variable groups using just VBA. One perfect example and our main factor to consider would be the Covid19 pandemic, this world crisis affected every single aspect of professional and daily life and activities and probably means that the return and total daily volume for each company is going to vary dramatically from what we could observe during 2017 and 2018.

The purpose of mentioning stock prices volatility is to put into context how important it is to ensure the best accuracy possible at least within the variables that we are considering in our analysis.

If we assume that the best case scenario to evaluate the return percentage would be that the total daily volume of every ticker equals 100% of our Total Volume for 2017 and 2018 dataset then let’s assume that our “accuracy” would be determined by the which percentage of the total daily volume every ticker represents.
2017                       |  2018
:-------------------------:|:-------------------------:
<img src="https://github.com/carloshgalvan95/stock-analysis/blob/main/Resources/Total_Volume_Percentage_2017.png" width="1500" />  |  <img src="https://github.com/carloshgalvan95/stock-analysis/blob/main/Resources/Total_Volume_Percentage_2018.png" width="1500" />

There is a pretty even distribution of the total volume percentage for all of our four possible tickers during 2018 so we are going to base our choice mainly using the total volume percentage of trades during 2017. What we can observe is that our safest two possible choices are **FSLR** and **SPWR**.

#### Staying green and accurate

For the future of the world, and Steve’s parents… the best road to take is staying on the greens. Considering accuracy and return, we would try make the safest best with the biggest possible return percentage of our investment.

Let’s say that we invested in the beginning of 2017, this is the return percentage that we would see at the end of 2018.

| Ticker | Return |        |           |
|--------|--------|--------|-----------|
|        | 2017   | 2018   | 2017-2018 |
| SPWR   | 23.1%  | -44.6% | -31.81%   |
| FSLR   | 101.3% | -39.7% | 21.36%    |
| ENPH   | 129.5% | 81.9%  | 317.55%   |
| RUN    | 5.5%   | 84.0%  | 94.154%   |

Based on this table we can clearly see that within our two possible investment options left there’s a clear winner. Our safest choice to ensure the highest probability of getting a return on our investment with the lowest risk possible would be **FSLR**.

But our options doesn’t end there, a second scenario and probably the better one if we have the necessary amount to invest and risk a little bit more would be to diversify our investment and go for both our safest bet and our biggest possible return investment, **FSLR** and **ENPH**.

