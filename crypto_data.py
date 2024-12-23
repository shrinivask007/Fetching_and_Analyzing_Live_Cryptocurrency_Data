import requests
import pandas as pd
import time
from openpyxl import Workbook

# Fetch live cryptocurrency data
def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1
    }
    response = requests.get(url, params=params)
    data = response.json()
    
    # Process data into a DataFrame
    df = pd.DataFrame(data, columns=[
        'name', 'symbol', 'current_price', 'market_cap',
        'total_volume', 'price_change_percentage_24h'
    ])
    return df

# Perform data analysis
def analyze_data(df):
    analysis = {}
    # Top 5 cryptocurrencies by market cap
    top_5 = df.nlargest(5, 'market_cap')[['name', 'market_cap']]
    analysis['Top 5 Cryptos'] = top_5
    
    # Average price of the top 50 cryptocurrencies
    avg_price = df['current_price'].mean()
    analysis['Average Price'] = avg_price
    
    # Highest and lowest 24-hour percentage price change
    highest_change = df.loc[df['price_change_percentage_24h'].idxmax()]
    lowest_change = df.loc[df['price_change_percentage_24h'].idxmin()]
    analysis['Highest Change'] = highest_change[['name', 'price_change_percentage_24h']]
    analysis['Lowest Change'] = lowest_change[['name', 'price_change_percentage_24h']]
    
    return analysis

# Save live data and analysis to Excel
def update_excel(df, analysis):
    filename = "crypto_live_data.xlsx"

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Write live data to the first sheet
        df.to_excel(writer, index=False, sheet_name="Live Data")
        
        # Write analysis results to the second sheet
        # 1. Top 5 Cryptos
        analysis['Top 5 Cryptos'].to_excel(writer, index=False, sheet_name="Analysis")
        
        # 2. Add Summary Info
        summary_data = {
            "Metric": ["Average Price", "Highest Change", "Lowest Change"],
            "Value": [
                f"${analysis['Average Price']:.2f}",
                f"{analysis['Highest Change']['name']} ({analysis['Highest Change']['price_change_percentage_24h']:.2f}%)",
                f"{analysis['Lowest Change']['name']} ({analysis['Lowest Change']['price_change_percentage_24h']:.2f}%)",
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, index=False, sheet_name="Analysis", startrow=7)

    print(f"Excel updated: {filename}")

# Main loop to fetch, analyze, and update Excel
def main():
    while True:
        try:
            crypto_data = fetch_crypto_data()
            analysis = analyze_data(crypto_data)
            update_excel(crypto_data, analysis)
            print("Data updated successfully.")
        except Exception as e:
            print("Error:", e)
        time.sleep(300)  # Update every 5 minutes

if __name__ == "__main__":
    main()
