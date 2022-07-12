# Securing a green future in Stocks

## Overview of the project

Steve’s parents are passionate about green energies, they believe that, as fossil fuels get used up, there will be more and more reliance on alternative energy production. However, Steve’s parents haven’t done much research and have instead decided to invest all their money into DAQO New Energy Corporation a company that makes silicon wafers for solar panels. Steve is concerned about diversifying their funds. He wants to analyze a handful of green energy stocks in addition to DAQO stock.

The main purpose of this analysis is to have more information and a point of comparison to base the decision of which company or companies to invest in on more than just a guess. We will be comparing eleven green energy companies based on three main indicators.

### Total Daily Volume of Trades

To ensure that the price that we are going to be evaluating reflects as accurately as possible the actual value of the stock we will be taking into consideration this first data point. Just as the central limit theorem states,

>*In probability theory, the* **central limit theorem (CLT)** *states that the distribution of a sample variable approximates a normal distribution (i.e., a “bell curve”) as the sample size becomes larger, assuming that all samples are identical in size, and regardless of the population's actual distribution shape.*
*-* *Ganti, A. (2022, July 8). What is the Central Limit Theorem (CLT)? Investopedia.*

The same idea applies to our stock value if the total daily volume of trades increases so as the accuracy of the value projections.

### Stocks in, stocks out

Our two other our two other data points are quite simple to define, the **stock value starting price** and the **ending price**, with this three data points we can get a general idea of how every company performed throughout the years that we will be using as our points of reference. For now, those years are going to be **2017** and **2018**.

## VBA as a powerful stock analysis tool

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

Now, how do we do this? First, we will define two variables to be able to evaluate how good our code actually works and how effective it is on reducing our run times, data can get quite extensive before we even realize, then we will be asking what year the user wants to run the analysis for:
```
Dim startTime As Single
Dim endTime  As Single
yearValue = InputBox("What year would you like to run the analysis on?")
```
