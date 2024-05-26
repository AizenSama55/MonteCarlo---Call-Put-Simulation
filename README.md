# MonteCarlo
Monte Carlo Stock Price Simulation
Welcome to the Monte Carlo Stock Price Simulation repository! This project offers a powerful and efficient way to simulate stock price movements using historical data and advanced financial indicators. Perfect for financial analysts, traders, and data scientists, this tool provides insights into the potential future prices of a stock and the probabilities of meeting specific target prices for options trading.

Features
Historical Data Integration: Load and analyze historical stock prices and volumes from Excel files to inform your simulations.
Financial Indicators Calculation: Utilize key indicators like RSI (Relative Strength Index) and MACD (Moving Average Convergence Divergence) to adjust simulations dynamically based on market conditions.
Monte Carlo Simulations: Run up to 1,000,000 simulations in parallel to forecast stock price movements with high accuracy.
Option Pricing: Evaluate the probability of meeting target prices for Call and Put options.
User-Friendly Interface: Interactively input parameters such as target price, option type, and days until expiry through a console-based interface.
Efficient Performance: Leverage multi-core processing to execute large-scale simulations swiftly.
Result Analysis: Calculate and display the probability of meeting the target price, and save the final simulated prices to a CSV file for further analysis.
How It Works
Input Parameters: Prompt the user to input the historical data file, option type (Call/Put), target price, and days until expiry.
Load Data: Read and process the historical stock price and volume data from the provided Excel file.
Calculate Indicators: Compute daily returns, mean return, volatility, RSI, and MACD to guide the simulation.
Run Simulations: Execute a large number of parallel Monte Carlo simulations, adjusting for market conditions using RSI and MACD.
Analyze Results: Determine the probability of the stock price meeting the target and output the results.
Save Output: Save the final simulated prices to a CSV file for detailed examination.
Installation
Clone the repository:

bash
Copy code
git clone https://github.com/yourusername/monte-carlo-stock-price-simulation.git
cd monte-carlo-stock-price-simulation
Ensure you have the necessary dependencies installed:

.NET Framework
ExcelDataReader
Build and run the project using your preferred IDE or the command line.

Usage
Run the executable or start the application from your IDE.
Follow the console prompts to input the historical data file name, option type, target price, and days until expiry.
View the probability results and check the CSV file for detailed final prices from the simulations.
Contributions
Contributions are welcome! If you have ideas for improvements or find any issues, feel free to open an issue or submit a pull request.
