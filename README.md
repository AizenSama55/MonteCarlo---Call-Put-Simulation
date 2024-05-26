Monte Carlo Simulation for Call and Put Options
This repository contains a Monte Carlo simulation for evaluating the probability of achieving a target price for call and put options. The simulation uses historical stock data to calculate daily returns, volatility, and various technical indicators such as RSI and MACD.

Features
Historical Data Loading: Fetches historical stock data using the YahooFinanceApi.
Daily Returns Calculation: Computes daily returns from the historical data.
Volatility Calculation: Calculates the volatility based on daily returns.
Technical Indicators: Includes calculations for Simple Moving Average (SMA), Exponential Moving Average (EMA), Relative Strength Index (RSI), and Moving Average Convergence Divergence (MACD).
Monte Carlo Simulation: Runs a large number of simulations to project future stock prices and evaluate the likelihood of meeting the target price for call or put options.
Parallel Processing: Utilizes parallel processing to improve the performance of the simulation.
CSV Export: Saves the final prices from the simulations to a CSV file.



git clone https://github.com/AizenSama55/MonteCarlo---Call-Put-Simulation.git
cd MonteCarlo---Call-Put-Simulation



Open in Visual Studio:
Open the solution file (MonteCarloSimulation.sln) in Visual Studio.

Run the Simulation:
Build and run the project. Follow the on-screen prompts to enter the stock symbol, option type (Call or Put), target price, and number of days until expiry.

Results:
The simulation will output the probability of meeting the target price based on the specified parameters and save the final prices to a CSV file (final_prices.csv).

Requirements
.NET Framework
YahooFinanceApi (version 2.3.3)

Installation
Ensure you have the required dependencies installed. You can install the YahooFinanceApi package via NuGet:
Install-Package YahooFinanceApi -Version 2.3.3


Contributing
Contributions are welcome! Please submit a pull request or open an issue to discuss any changes or improvements.
