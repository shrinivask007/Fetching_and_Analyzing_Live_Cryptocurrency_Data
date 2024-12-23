import requests
import pandas as pd
import time

# Fetch live data from CoinGecko API
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

# Save to Excel
def update_excel(df):
    filename = "crypto_live_data.xlsx"
    with pd.ExcelWriter(filename, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, index=False, sheet_name="Top 50 Cryptos")
    print("Excel updated:", filename)

# Main loop to fetch and update data
while True:
    try:
        crypto_data = fetch_crypto_data()
        update_excel(crypto_data)
        time.sleep(300)  # Update every 5 minutes
    except Exception as e:
        print("Error:", e)
