import requests
import pandas as pd
import io
import zipfile
from datetime import datetime, timedelta
from io import StringIO
from requests.exceptions import RequestException
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
import random
import os

timeout = 30
server_overloaded = False

# Define series priority
series_priority = {' EQ': 1, ' BE': 2}
default_priority = 3



def get_random_headers():
    user_agents = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36",
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36",
        # Add more user agents as needed
    ]
    headers = {
        "User-Agent": random.choice(user_agents),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.5",
        "Connection": "keep-alive",
    }
    return headers



def wait_between_requests(min_delay=5, max_delay=15):
    delay = random.uniform(min_delay, max_delay)
    print(f"Waiting for {delay:.2f} seconds before retrying...")
    time.sleep(delay)


proxies = [
    {"http": "http://proxy1.com:8080", "https": "https://proxy1.com:8080"},
    {"http": "http://proxy2.com:8080", "https": "https://proxy2.com:8080"},
    # Add more proxies
]

def get_random_proxy():
    return random.choice(proxies)


def get_bse_bhavcopy_url(date_input):
    """
    Returns the BSE Bhavcopy URL based on the input date.
    
    Parameters:
    date_input (datetime): The date for which the URL is to be generated.

    Returns:
    str: The URL for the BSE Bhavcopy.
    """
    if date_input < datetime(2024, 1, 1):
        formatted_date = date_input.strftime('%d%m%Y')
        return f"http://www.bseindia.com/download/BhavCopy/Equity/BSE_EQ_BHAVCOPY_{formatted_date}.zip"
    else:
        formatted_date = date_input.strftime('%Y%m%d')
        return f"https://www.bseindia.com/download/BhavCopy/Equity/BhavCopy_BSE_CM_0_0_0_{formatted_date}_F_0000.csv"

def get_nse_bhavcopy_url(date_input):
    """
    Returns the NSE Bhavcopy URL based on the input date.

    Parameters:
    date_input (datetime): The date for which the URL is to be generated.

    Returns:
    str: The URL for the NSE Bhavcopy.
    """
    dd = date_input.strftime('%d')
    mm = date_input.strftime('%m')
    yyyy = date_input.strftime('%Y')
    return f"https://nsearchives.nseindia.com/products/content/sec_bhavdata_full_{dd}{mm}{yyyy}.csv"


def fetch_nse_bhavcopy(date_input, retries=3):
    """
    Fetches the NSE Bhavcopy for the given date and returns only the SYMBOL and CLOSE_PRICE columns.

    Parameters:
    date_input (str or datetime): The date for which the Bhavcopy is to be fetched.
    retries (int): Number of retry attempts in case of failure.

    Returns:
    pd.DataFrame or None: DataFrame with SYMBOL and CLOSE_PRICE columns, or None if fetching fails.
    """
    if isinstance(date_input, str):
        date_input = datetime.strptime(date_input, "%Y-%m-%d")

    directory = "./bhavcopies"
    os.makedirs(directory, exist_ok=True)
    filename = os.path.join(directory, f"nse_bhavcopy_{date_input.strftime('%d_%m_%Y')}.csv")

    if os.path.exists(filename):
        print(f"File {filename} already exists. Loading from local storage.")
        data = pd.read_csv(filename)
        data['SERIES_PRIORITY'] = data[' SERIES'].map(series_priority).fillna(default_priority)
        filtered_data = data.drop_duplicates(subset='SYMBOL', keep='first')
        return filtered_data[["SYMBOL", " CLOSE_PRICE"]]

    url = get_nse_bhavcopy_url(date_input)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }

    retry_delay = 30  # Initial delay in seconds
    for attempt in range(retries):
        try:
            response = requests.get(url, headers=headers, timeout=30)  # Increased timeout
            response.raise_for_status()
            data = pd.read_csv(io.StringIO(response.content.decode('utf-8')))

            # Save the full Bhavcopy to file
            data.to_csv(filename, index=False)
            print(f"Saved Bhavcopy to {filename}")
            data['SERIES_PRIORITY'] = data[' SERIES'].map(series_priority).fillna(default_priority)
            filtered_data = data.drop_duplicates(subset='SYMBOL', keep='first')
            return filtered_data[["SYMBOL", " CLOSE_PRICE"]]

        except RequestException as e:
            print(f"Attempt {attempt + 1} failed: {date_input}")
            if response and (response.status_code in {503, 429}):
                print(f"Rate limit exceeded. Pausing for {retry_delay} seconds...")
                time.sleep(retry_delay)
                retry_delay *= 2  # Exponentially increase the delay
                retry_delay += random.uniform(0, 1)
            else:
                return None
            if attempt < retries - 1:
                print("Retrying...")

    print("Failed to fetch NSE Bhavcopy after multiple attempts.")
    return None

def fetch_bse_bhavcopy(date_input, retries=3):
    if isinstance(date_input, str):
        date_input = datetime.strptime(date_input, "%Y-%m-%d")

    directory = "./bhavcopies"
    os.makedirs(directory, exist_ok=True)
    filename = os.path.join(directory, f"bse_bhavcopy_{date_input.strftime('%d_%m_%Y')}.csv")

    if os.path.exists(filename):
        print(f"File {filename} already exists. Loading from local storage.")
        data = pd.read_csv(filename)
        return data[["TckrSymb", "ClsPric"]]

    url = get_bse_bhavcopy_url(date_input)

    retry_delay = 30
    for attempt in range(retries):
        try:
            headers = get_random_headers()
            response = requests.get(url, headers=headers, timeout=30)
            response.raise_for_status()

            if date_input < datetime(2024, 1, 1):
                with zipfile.ZipFile(io.BytesIO(response.content)) as zf:
                    csv_file_name = zf.namelist()[0]
                    with zf.open(csv_file_name) as file:
                        data = pd.read_csv(file)
            else:
                data = pd.read_csv(io.StringIO(response.content.decode("utf-8")))

            data.to_csv(filename, index=False)
            print(f"Saved Bhavcopy to {filename}")
            return data[["TckrSymb", "ClsPric"]]

        except requests.exceptions.RequestException as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            if attempt < retries - 1:
                wait_between_requests()
                retry_delay *= 2
            else:
                break

    print("Failed to fetch BSE Bhavcopy after multiple attempts.")
    return None

def fetch_data_for_ticker(ticker, stock_name, date_input, nse_data, bse_data):
    symbol, exchange = ticker.split(".")
    close_price = None

    if exchange.lower() == "ns" and nse_data is not None:
        row = nse_data[nse_data["SYMBOL"].str.strip() == symbol]
        if not row.empty:
            close_price = row.iloc[0][" CLOSE_PRICE"]
    
    elif exchange.lower() == "bo" and bse_data is not None:
        if date_input < datetime(2024, 1, 1):
            row = bse_data[bse_data["SC_NAME"].str.strip() == stock_name.strip()]
            if not row.empty:
                close_price = row.iloc[0]["CLOSE"]
        else:
            row = bse_data[bse_data["TckrSymb"].str.strip() == symbol]
            if not row.empty:
                close_price = row.iloc[0]["ClsPric"]
    
    return close_price

 
def get_stock_data(date_input, keys):
    """Fetches the close prices for the given tickers based on their exchange (NSE or BSE)."""
    try:
        if isinstance(date_input, str):
            date_input = datetime.strptime(date_input, "%Y-%m-%d")
        elif not isinstance(date_input, datetime):
            date_input = datetime.fromtimestamp(date_input.timestamp())
        
        nse_data = fetch_nse_bhavcopy(date_input)
        bse_data = fetch_bse_bhavcopy(date_input)
        
        if nse_data is None or bse_data is None:
            return None
        
        if nse_data is None:
            print("No close prices for nse")
            print(bse_data)
            
        if bse_data is None:
            print("No close prices for bse")
            print(nse_data)
        
        nse_data = nse_data.rename(columns={"SYMBOL": "Symbol", " CLOSE_PRICE": "CLOSE_PRICE"})
        nse_data["Symbol"] = nse_data["Symbol"] + ".NS"
        bse_data = bse_data.rename(columns={"TckrSymb": "Symbol", "ClsPric": "CLOSE_PRICE"})
        bse_data["Symbol"] = bse_data["Symbol"] + ".BO"

        # Concatenating both dataframes
        merged_data = pd.concat([nse_data, bse_data], ignore_index=True)

        # Removing duplicates, if any
        merged_data = merged_data.drop_duplicates(subset=["Symbol", "CLOSE_PRICE"])


        # Debug: Check for duplicate labels

        # Select symbols based on the date
        symbols_to_fetch = keys

        # close_prices = merged_data.set_index("Symbol")[["CLOSE_PRICE"]].T[symbols_to_fetch].values
        # close_prices = merged_data.set_index("Symbol").drop(["Symbol"], axis=1)[["CLOSE_PRICE"]].reindex(symbols_to_fetch).T.values
        close_prices = merged_data.set_index("Symbol")["CLOSE_PRICE"]
        if not close_prices.index.is_unique:
            print("not unique index")
            duplicates = merged_data[["Symbol", "CLOSE_PRICE"]][merged_data["Symbol"].duplicated()]
            print("Duplicate Symbols:", duplicates)
            close_prices = close_prices.groupby(close_prices.index).first()
       
        close_prices = close_prices.reindex(symbols_to_fetch)

        return close_prices
    except Exception as e:
        print(f"Failed to fetch close prices: {e}  {date_input}")
        return None

def create_stock_price_df(start_date, end_date, keys, symbols_dict):
    """Create a DataFrame with stock close prices for a range of dates."""
    start_date = datetime.strptime(start_date, "%Y-%m-%d")
    end_date = datetime.strptime(end_date, "%Y-%m-%d")
    
    dates = pd.date_range(start=start_date, end=end_date).to_pydatetime().tolist()
    dates = [date for date in dates if date.weekday() != 6]  # Skip Sundays
 
    def fetch_data_for_date(date):
        if server_overloaded:
            time.sleep(30)
        date_str = date.strftime("%Y-%m-%d")
        return get_stock_data(date_str, symbols_dict[date_str])
    
    with ThreadPoolExecutor() as executor:
        results = list(executor.map(fetch_data_for_date, dates))
    
    # Filter out None results and their corresponding dates
    filtered_data = [(d.strftime("%Y-%m-%d"), r) for d, r in zip(dates, results) if r is not None]

    # Extract the filtered dates and results
    filtered_dates = [item[0] for item in filtered_data]
    filtered_results = [item[1] for item in filtered_data]

    print("Filtered results")
    print(filtered_results)
    # Create the DataFrame
    # price_df = pd.DataFrame(filtered_results, index=filtered_dates, columns=keys)
    price_df = pd.DataFrame([series.values for series in filtered_results], index=filtered_dates, columns=keys)
    # Reset the index to make dates a column
    price_df = price_df.T.reset_index()

    # Rename the columns for clarity
    price_df = price_df.rename(columns={"index": "ticker"})
    print("shape", price_df.shape)
    return price_df
