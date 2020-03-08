# stock-analysis
This analysis of 12 Wallstreet stocks aims to determine annual performance of each stock in 2017 and 2018 for the purpose of gaining a better understanding of whether investing in a particular stock would be a good decision.

## stock-analysis.xlsm file
### DQ Analysis
To begin this work, the subroutine "DQ Analysis" was written in VBA to analyze the stock, DAQO (DQ), for the year 2018. Total Daily Volume and Return were calculated.

The data indicates poor performance in 2018, showing a 62.6% loss, as summarized in the table below:
| Year | Total Daily Volume | Return |
|------|--------------------|--------|
| 2018 | 107,873,900 | -62.6%|

### All Stocks Analysis
The analysis of DQ was expanded upon so that all 12 stocks were included in the analysis for both years (2017 and 2018). The subroutine *AllStocksAnalysis()* was written in VBA to perform the analysis. Additionally, the subroutine *ClearWorksheet()* was written to clear the analysis from the output worksheet, **All Stocks Analysis**.

#### Explanation of VBA
The macro *AllStocksAnalysis()* involves the following variables:
* **yearValue** (values = "2017" and "2018")
* **tickers()** (Array of 12 tickers defined 'As String')
* **startingPrice** (defined 'As Single')
* **endingPrice** (defined 'As Single')
* **ticker** (initialized as equal to the current ticker, tickers(i))
* **totalVolume** (initialized as equal to zero so that it can be used in a sum)
* **RowCount** (defined by formula to calculate number of rows in worksheet)
* **dataRowStart** (created for the output sheet for the purposes of applying conditional formatting)
* **dataRowEnd** (created for the output sheet for the purposes of applying conditional formatting)

The subroutine is broken down in the following steps:
1. The user types the year upon which to perform the analysis (i.e. the user defines yearValue as "2017" or "2018").
2. On the output worksheet "All Stocks Analysis" the following is defined:
    * The cell A1 value = "All Stocks (+ "yearValue" +)"
    * The cells A3, B3, and C3  represent the header row with the headings "Ticker", "Total Daily Volume", and "Return", respectively.
3. The array tickers(12) is defined as a string.
    * values = "AY", "CSIQ", "DQ", "ENPH", "FLSLR", "HASI", "JKS", "RUN", "SEDG", SPWR", "TERP", and "VSLR"
4. The variables startingPrice and endingPrice are initialized As Single.
5. The code activates the data worksheet upon which to run the analysis as defined by yearValue ("2017" or "2018").
6. The code determines the number of rows in the active data sheet upon which to run the analysis (using formula for calculating RowCount).
7. The code runs through the data for the array of tickers(12), and totalVolume is initialized as equal to zero.
    - The code loops through rows 2 to RowCount for each ticker()
        - **totalVolume** for each ticker() is defined by adding the value for volume (in column H of the raw data) when the value for "Ticker" (in column A of the raw data) is equal to the **ticker** (i.e. *tickers(i)*) value.
            - This is achieved by using a condition.
            - The value is able to be calculated since **totalVolume** was initialized as equal to zero before running the loop.
        - **startingPrice** is determined by taking the value for "Close" (in column F of raw data) when the value for "Ticker" (in column A of the raw data) in the current row is different than the value for "Ticker" (in column A of the raw data) in the previous row.
            - This is achieved by using a condition.
        - **endingPrice** is determined by taking the value for "Close" (in column F of raw data) when the value for "Ticker" (in column A of the raw data) in the current row is different than the value for "Ticker" (in column A of the raw data) in the next row.
            - This is achieved by using a condition.
8. The code populates the output for each ticker on the activated worksheet "All Stocks Analysis" as follows:
    - Under the heading "Ticker", populate the value for **ticker**
    - Under the heading "Total Daily Volume", populate the value for **totalVolume**
    - Under the heading "Return", populate the value for the formula **endingPrice** / **startingPrice** - 1

9. The output sheet is formatted as follows:
    - Number format for Total Daily Volume (**totalVolume**) is defined as "#,##0"
    - Number format for Return is defined as "0.00%" (percentage to the nearest hundreth)
    - Conditional formatting is applied to Return column from dataRowStart to dataRowEnd such that if the value is greater than zero, then the cell will be green; if the value is less than zero, then the cell will be red; otherwise, the cell will not be coloured.

#### Output

To run 

    NOTE: additional formatting for font and borders was applied which has not been menitoned in step 9

## Methodology
