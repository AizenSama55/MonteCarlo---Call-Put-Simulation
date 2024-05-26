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
