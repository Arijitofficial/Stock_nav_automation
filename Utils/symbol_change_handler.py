import pandas as pd
from datetime import datetime, timedelta

file_path = 'Excels\symbolchange.csv'
df = pd.read_csv(file_path, encoding="latin1", header=None)
df.columns = ["name", "old_symbol", "new_symbol", "date"]
df = df[df["old_symbol"] != df["new_symbol"]]

def map_symbols(symbol_list, start_date, end_date):
    """
    Maps symbols based on a given date range and a dataframe of symbol changes.

    Args:
        symbol_list (list): List of stock symbols to process.
        df (pd.DataFrame): DataFrame containing columns ["name", "old_symbol", "new_symbol", "date"].
        start_date (str): Start date of the range (format: 'YYYY-MM-DD').
        end_date (str): End date of the range (format: 'YYYY-MM-DD').

    Returns:
        dict: A dictionary where the key is the date and the value is the updated symbol list for that date.
    """
    # Convert start_date and end_date to datetime objects
    start_date = datetime.strptime(start_date, "%Y-%m-%d")
    end_date = datetime.strptime(end_date, "%Y-%m-%d")

    # Ensure the date column in df is in datetime format
    df["date"] = pd.to_datetime(df["date"])

    # Determine the maximum date in the dataframe
    max_date_in_df = df["date"].max()

    # Determine the starting date for the loop
    loop_start_date = max(max_date_in_df, end_date)

    # Initialize the result dictionary
    date_to_symbols = {}

    # Create a copy of the symbol list to modify
    updated_symbols = symbol_list[:]

    # Iterate from loop_start_date to start_date in reverse
    current_date = loop_start_date
    while current_date >= start_date:
        # Filter the DataFrame for changes effective on the current_date
        daily_changes = df[df["date"] == current_date]


        # Add the updated symbols to the result dictionary if the current_date is within the range
        if start_date <= current_date <= end_date:
            date_to_symbols[current_date.strftime("%Y-%m-%d")] = updated_symbols[:]
            
        # Apply changes from the daily_changes DataFrame
        for _, row in daily_changes.iterrows():
            updated_symbols = [
                row["old_symbol"]+'.NS' if symbol == (row["new_symbol"]+'.NS') else symbol
                for symbol in updated_symbols
            ]

        

        # Move to the previous day
        current_date -= timedelta(days=1)

    return date_to_symbols
