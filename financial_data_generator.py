import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta

"""all tickers work, need to test excel, see if i can get pyqtgraph etc to automatically do pricing graphs as well"""
#ticker names 
symbols = [
    "HYZN", "NRGV", "HYMTF", "TM", "NKLA",
    "CMI", "0AB9.L", "HYLN", "BLDP", "PLUG"
]

# gets dates for now , 1 week ago , 1 month ago 
end_date = datetime.now()
start_date_week = end_date - timedelta(weeks=1)
start_date_month = end_date - timedelta(weeks=4)

# init empty list for data
data = []

for symbol in symbols:
    stock = yf.Ticker(symbol)

    # Fetch historical market data
    hist = stock.history(period="1mo")
    
    # Check if data is available
    if not hist.empty:
        # current price 
        current_price = hist['Close'].iloc[-1]
        
        # prices 1 week, 1 month ago
        price_week_ago = hist.loc[hist.index >= start_date_week.strftime('%Y-%m-%d')]['Close'].iloc[0] if not hist.loc[hist.index >= start_date_week.strftime('%Y-%m-%d')].empty else None
        price_month_ago = hist.loc[hist.index >= start_date_month.strftime('%Y-%m-%d')]['Close'].iloc[0] if not hist.loc[hist.index >= start_date_month.strftime('%Y-%m-%d')].empty else None
        
        # % change in price, handling cases where data might be missing
        change_week = ((current_price - price_week_ago) / price_week_ago) * 100 if price_week_ago else None
        change_month = ((current_price - price_month_ago) / price_month_ago) * 100 if price_month_ago else None
        
        # market cap and no of shares 
        market_cap_m = stock.info.get('marketCap', 0) / 1e6 if stock.info.get('marketCap') else None
        shares_outstanding = stock.info.get('sharesOutstanding', 0)
        
        # full company name
        full_company_name = stock.info.get('longName', 'Name not found')
        
        # append these metrics to the list - for now set primary industry to na cause this is only getting financial data from yfinance not qualitative 
        data.append({
            "Company Name": full_company_name,
            "Ticker:": symbol,
            "Primary Industry": "N/A",  # qualitative , find a way to fill in later 
            "Current Price": current_price,
            "Price 1 Week Ago": price_week_ago,
            "Price 1 Month Ago": price_month_ago,
            "Change in Price (%) Week on Week": change_week,
            "Change in Price (%) 1 Month Ago": change_month,
            "Market Cap (M) Current": market_cap_m,
            "Number of Shares Current": shares_outstanding
        })
    else:
        print(f"No data found for {symbol}, it may be delisted or unavailable.")

# create a dataframe from the data list (using to export to csv )
df = pd.DataFrame(data)

# export to csv 
csv_file_name = "company_data.csv"
df.to_csv(csv_file_name, index=False)
print(f"Data exported to {csv_file_name}")

