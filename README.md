Overview
The script performs a Monte Carlo simulation to predict stock price movements based on historical data. It calculates the probability of a stock reaching a certain target price by a specified expiry date, considering the option type (Call or Put). The script integrates financial indicators such as RSI (Relative Strength Index) and MACD (Moving Average Convergence Divergence) to adjust the simulations dynamically.

Detailed Explanation
Initialization and Setup:

The script begins by registering the encoding provider to handle different text encodings.
It enters a continuous loop to allow the user to run multiple simulations without restarting the console application.
User Input:

File Name: The user is prompted to enter the filename of the historical data Excel file.
Option Type: The user specifies whether the option is a Call or Put.
Target Price: The user inputs the strike price they are interested in.
Days Until Expiry: The user specifies the number of days until the option expires.
Data Loading:

The script attempts to load historical stock prices and volumes from the specified Excel file.
It reads the data from the first sheet of the Excel file, assuming that the "Close" prices are in the fifth column and volume data is in the sixth column.
If the data is successfully loaded, it extracts the prices into a list.
Calculate Daily Returns and Volatility:

Daily Returns: The script calculates the daily returns as the percentage change in price from one day to the next.
Mean Return: The average of the daily returns.
Volatility: The standard deviation of the daily returns, representing the stock's price fluctuation over time.
Simulation Parameters:

The script sets the current price to the last price in the historical data.
It defines the number of simulations to run (1,000,000).
Monte Carlo Simulations:

Parallel Processing: Uses multi-core processing to run simulations in parallel for efficiency.
Random Number Generation: Uses a thread-safe random number generator.
Simulation Loop: For each simulation, the script:
Initializes the simulated price to the current price.
Calculates RSI and MACD based on historical prices.
Applies a random shock to the price, adjusted for volatility and market conditions (overbought/oversold as indicated by RSI).
Updates the simulated price and appends it to the historical prices for the next day's calculation.
Result Analysis:

Target Count: Counts the number of simulations where the final price meets or exceeds the target price for Call options or falls below the target price for Put options.
Probability Calculation: Calculates the probability of the target price being met by dividing the target count by the total number of simulations.
Output Results:

The script prints the number of successful simulations and the probability of meeting the target price.
It saves the final prices from all simulations to a CSV file for further analysis.
Functions
LoadHistoricalData(filePath As String) As List(Of (Double, Double)):

Loads historical price and volume data from the specified Excel file.
Returns a list of tuples containing prices and volumes.
CalculateDailyReturns(prices As List(Of Double)) As List(Of Double):

Calculates and returns the daily returns based on the provided prices.
CalculateVolatility(returns As List(Of Double)) As Double:

Calculates and returns the volatility as the standard deviation of the daily returns.
CalculateSMA(prices As List(Of Double), period As Integer) As Double:

Calculates and returns the Simple Moving Average (SMA) over the specified period.
CalculateEMA(prices As List(Of Double), period As Integer) As Double:

Calculates and returns the Exponential Moving Average (EMA) over the specified period.
CalculateRSI(prices As List(Of Double), period As Integer) As Double:

Calculates and returns the Relative Strength Index (RSI) over the specified period.
CalculateMACD(prices As List(Of Double), shortPeriod As Integer, longPeriod As Integer, signalPeriod As Integer) As Double:

Calculates and returns the Moving Average Convergence Divergence (MACD) and the MACD signal line.
SaveFinalPricesToFile(finalPrices As Double(), filePath As String):

Saves the final prices from all simulations to a CSV file.
Summary
The script uses a combination of historical data analysis, financial indicators, and Monte Carlo simulation techniques to predict stock price movements and evaluate the probability of meeting target prices for options. It leverages parallel processing for efficiency and allows continuous user interaction for multiple simulations. The results provide valuable insights for traders and analysts in making informed decisions about option pricing and stock movements.


How to Use This Monte Carlo Stock Price Simulation
Prerequisites
Download Historical Data:

Go to Yahoo Finance.
Search for the stock you are interested in.
Click on "Historical Data" and download the data in CSV format.
Convert to Excel (.xlsx) Format:

Open the downloaded CSV file in Excel.
Save the file as an Excel Workbook (.xlsx).
Place the Excel File:

Place the .xlsx file in the same directory as the Monte Carlo Simulation program executable.
Running the Program
Start the Program:

Run the Monte Carlo Simulation executable or start the application from your IDE.
Follow the Prompts:

Enter the filename for historical data: Type the name of the Excel file containing the historical data (e.g., historical_data.xlsx) and press Enter.
Call or Put option: Type Call for a Call option or Put for a Put option and press Enter.
Enter the target price: Type the strike price (target price) and press Enter.
Enter the number of days until expiry: Type the number of days until the option expires and press Enter.
View the Results:

The program will run the simulations and display the probability of the stock price meeting the target price.
The final prices from all simulations will be saved to a file named final_prices.csv in the same directory.
Example Usage
Download Historical Data:

Search for "AAPL" on Yahoo Finance.
Download the historical data as a CSV file.
Convert to Excel Format:

Open the downloaded CSV file in Excel.
Save it as aapl_historical_data.xlsx.
Run the Program:

Place aapl_historical_data.xlsx in the program directory.
Start the program and follow the prompts:
mathematica
Copy code
Enter the filename for historical data: aapl_historical_data.xlsx
Is this a Call or Put option? (Enter 'Call' or 'Put'): Call
Enter the target price: 150
Enter the number of days until expiry: 30
View the results in the console and check final_prices.csv for the detailed output.
Notes
Ensure that the Excel file is correctly formatted with historical prices in the appropriate columns (e.g., "Close" prices in the fifth column).
If you encounter any errors, check the console for detailed error messages and ensure that the file paths and input values are correct.
