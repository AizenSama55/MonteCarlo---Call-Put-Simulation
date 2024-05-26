import matplotlib.pyplot as plt
import csv

def load_final_prices(file_path):
    final_prices = []
    with open(file_path, 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            final_prices.append(float(row[0]))
    return final_prices

def visualize_results(final_prices, target_price):
    plt.hist(final_prices, bins=50, alpha=0.75)
    plt.axvline(x=target_price, color='r', linestyle='--')
    plt.xlabel('Final Price')
    plt.ylabel('Frequency')
    plt.title('Monte Carlo Simulation Results')
    plt.show()

if __name__ == "__main__":
    final_prices = load_final_prices('final_prices.csv')
    target_price = 240
    visualize_results(final_prices, target_price)
