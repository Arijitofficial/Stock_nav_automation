import numpy as np
import pandas as pd
from datetime import datetime, timedelta
from dateutil.parser import parse
from Utils.sales_purchase_util import init_dict, save_sales_purchase_dict
from Utils.drill_down_util import enter_track, init_drill_down_df, save_drill_down_df
from down_close_price_data import get_stock_data, create_stock_price_df
from Utils.symbol_change_handler import map_symbols




# Define the file containing the records
# C:\Users\ASUS\Documents\automation\Excels\Shares as on 18.01.2025.xlsx
file_path = './Excels/Shares as on 18.01.2025.xlsx'
sheet_name = "Sheet1"

# Read the data from the Excel file if it exists or take file path input

try:

    df = pd.read_excel(file_path, sheet_name, engine='openpyxl')

except Exception as e:

    print("File not found. Please provide a valid file path.")
    print("make sure the file is present in the same directory as the code file")
    file_path = input("enter the file name: ")
    sheet_name = input("enter the sheet name: ")

    df = pd.read_excel(file_path, sheet_name, engine='openpyxl')



df = df[(df["Cat"].str.lower() == 'normal') | (df["Cat"].str.lower() == 'others')]

# Fill null entries in the "NSE Name " column with the corresponding non-null "Name of Shares" entries
df["NSE Name "] = df.groupby("Name of Shares")["NSE Name "].transform(lambda x: x.fillna(method='ffill').fillna(method='bfill'))


df["Broker"].fillna(value="unknown", inplace=True)
df["Broker"] = df["Broker"].str.replace(r"(?i)\bone\b", "IIFL", regex=True)
# Verify the result
print("null count of symbol", df["NSE Name "].isnull().sum())
print(df[df["NSE Name "].isnull()]["Name of Shares"].unique())

print(df["Broker"].unique())


sales_purchase_dict = init_dict(broker_names=df['Broker'].unique())
drill_down_df = init_drill_down_df()

# input start and end date and validate the string inputs in yyyy-mm-dd format
while True:
    if "Overall" in sales_purchase_dict and len(sales_purchase_dict["Overall"]) > 0:
        print("Enter 1 to pick the start date from sales purchase records\nEnter 2 to pick the start date manually", end="\nput choice: ")
        choice = input()
    else:
        choice = "2"

    if choice == "1":
        if "Overall" in sales_purchase_dict and len(sales_purchase_dict["Overall"]) > 0:
            start_date = str(sales_purchase_dict["Overall"]['Date'].iloc[-1])
        else:
            # Take date input if 'Overall' is not found or has no rows
            print("Sales Purchase track does not exist")
            start_date = input("Enter start date(yyyy-mm-dd): ")
        pass
    else:
        start_date = input("Enter start date(yyyy-mm-dd): ")

    end_date = input("Enter end date(yyyy-mm-dd): ")

    try:
        start_date = parse(start_date).date()
        end_date = parse(end_date).date()
        break

    except Exception as e:
        print(e)
        print("please enter valid date in yyyy-mm-dd format")


# keep timestamp
start_timestamp = datetime.now()


# ----------------------------------------------------------------
#  close price fetching for all stocks
# ----------------------------------------------------------------


# Convert 'S. Date' to datetime, if it exists in the DataFrame
df['S. Date'] = pd.to_datetime(df['S. Date'], errors='coerce').dt.date

df.to_csv('unfiltered_df.csv')
# Filter rows based on sell_date
filtered_df = df[
    (df['S. Date'].isnull()) | (df['S. Date'] > start_date)
]
filtered_df.to_csv('filtered_df.csv')


# Select unique 'NSE Name' and 'Symbol' values and add 'ticker' column
unique_df = filtered_df[['NSE Name ', 'Symbol', 'Name of Shares']].drop_duplicates()

unique_df['ticker'] = unique_df.apply(
    lambda row: f"{row['NSE Name ']}.NS" if row['Symbol'] == "NSE" else f"{row['NSE Name ']}.BO", 
    axis=1
)

unique_df['bhav_keys'] = unique_df.apply(
    lambda row: f"{row['NSE Name ']}.NS" if row['Symbol'] == "NSE" else f"{row['Name of Shares']}.BO", 
    axis=1
)

unique_df.to_csv('unique_df.csv')
print("BSE data", unique_df[unique_df['Symbol']=='BSE'])


keys = list(set(unique_df['ticker']))
bhav_keys = list(set(unique_df['bhav_keys']))



# tstamp = datetime.now()
# Initialize an empty DataFrame to store the data
# all_data = asyncio.run(process_tickers(unique_df, start_date, end_date))

# Loop through each unique ticker

# # Loop through each unique ticker
# for row_ind, row in unique_df.iterrows():
#     nse_name = row["NSE Name "]
#     ticker = row["ticker"]

#     try:
       

#         # Append to the main DataFrame
#         # if not data.isnull().all(axis=1).any():
#         #     all_data = pd.concat([all_data, data], axis=0)
#         if ticker.endswith(".NS"):
#             nse_data = fetch_historical_stock_data(nse_name, start_date, end_date)
#             # if not data.isnull().all(axis=1).any():
#             all_data = pd.concat([all_data, nse_data], axis=0)
#             print("nse data")
#             print(nse_data)
#         else:
#              # Download stock data
#             bse_data = yf.download(ticker, start=start_date, end=(end_date + timedelta(days=1)))

#             # Select and rename the 'Close' column
#             bse_data = bse_data[['Close']]
#             bse_data.columns = [nse_name]

#             # Forward-fill missing values
#             bse_data = bse_data.ffill()
#             bse_data.index = bse_data.index.strftime("%Y-%m-%d")

#             # Transpose the bse_dataFrame
#             bse_data = bse_data.T
#             bse_data.index.name = "Ticker"
#             # if not data.isnull().all(axis=1).any():
#             all_data = pd.concat([all_data, bse_data], axis=0)

#         # # Export after processing at least two tickers
#         # if int(row_ind) > 1:
#         #     print("Index", row_ind)
#         #     print(all_data)
#         #     all_data.to_csv("yahoo_many_data.csv")
#         #     break

#     except Exception as e:
#         print(f"Could not retrieve data for {ticker}: {e}")


# Reset the index to make "NSE Name" a column
# all_data = all_data.reset_index()
# print(datetime.now() - tstamp)

# Save the data to CSV in the desired format
# all_data.to_csv("nse_close_price_data.csv", index=False)
# print("Data saved to nse_close_price_data.csv\n\n\n\n")
tstamp = datetime.now()
iter_date = start_date
# while iter_date < end_date:
#     if iter_date.weekday() == 6:
#         iter_date += timedelta(days=1)
#         continue
#     get_stock_data(str(iter_date), unique_df['ticker'], unique_df['Name of Shares'])
#     iter_date += timedelta(days=1)
print("bhavcopies getting fetched")

symbols_dict = map_symbols(keys, start_date=str(start_date), end_date=str(end_date))
close_price_df = create_stock_price_df(start_date=str(start_date), end_date=str(end_date), keys = keys, symbols_dict=symbols_dict)
close_price_df.to_csv("close_prices_data_2nd_method.csv", index=False)


print(datetime.now() - tstamp)

cp_file_path = "./Excels/nse_close_price_data.csv"
# close_price_df = all_data

# print("columns", close_price_df.columns)
# raise ValueError("artificial error")

# iterate over start to end dates and fetch current market prices and update files

# Initialize current_date
current_date = start_date
temp_df = df.copy()
unlisted_shares = {}
while current_date <= end_date:
    print(f"Processing date: {current_date}")

    # Create a temporary DataFrame to store data for the current date
    
    total_market_value = 0

    values = {"Overall": 0}
    for broker_name in df['Broker'].unique():

        values[broker_name] = 0

    sales = {"Overall": 0}
    for broker_name in df['Broker'].unique():
        sales[broker_name] = 0

    purchase = {"Overall": 0}
    for broker_name in df['Broker'].unique():
        purchase[broker_name] = 0

    net_fund = {"Overall": 0}
    for broker_name in df['Broker'].unique():
        net_fund[broker_name] = 0

    units = {"Overall": 0}
    for broker_name in df['Broker'].unique():
        units[broker_name] = 0

    nav = {"Overall": 0}
    for broker_name in df['Broker'].unique():
        nav[broker_name] = 0

    for index, row in temp_df.iterrows():
        symbol = row['NSE Name ']
        dop = pd.to_datetime(row['DOP'])
        sell_date = pd.to_datetime(row['S. Date']) if pd.notnull(row['S. Date']) else None

        # Check if stock was available on the current date based on DOP and S. Date
        if pd.notnull(symbol) and (dop <= pd.to_datetime(current_date)) and (sell_date is None or pd.to_datetime(current_date) <= sell_date):
            symbol_ns = symbol + (".NS" if row['Symbol'] == "NSE" else ".BO")
            try:
                # Check if the symbol exists in the 'Price' column
                if symbol_ns in close_price_df['ticker'].values:
                    # Check if the date column exists in close_price_df
                    date_column = current_date.strftime("%Y-%m-%d")

                    # if the market was open on the current date
                    if date_column in close_price_df.columns:
                        close_price = close_price_df.loc[close_price_df['ticker'] == symbol_ns, date_column].values[0]

                        if pd.isna(close_price) or pd.isnull(close_price):
                            close_price = row['Cost/Sh']
                            # unlisted_shares.add(symbol)
                    else:
                        # Raise an error if the date column is not found
                        raise ValueError(f"Date {date_column} not found in the data.")
                else:
                    print("Stock symbol not found in the provided data.")
                    close_price = None

            except Exception as e:
                print(f"Error fetching data for {symbol_ns} on {current_date}: {e}")
                date = pd.to_datetime(current_date)
                if (date.dayofweek == 5) or (date.dayofweek == 6):
                    print("weekend day", current_date)

                # Iterate through each DataFrame in the dictionary
                if len(sales_purchase_dict["Overall"]) <= 0:
                    break
                for key, sp_df in sales_purchase_dict.items():
                    # Ensure date column is in datetime format
                    if (len(sp_df) <= 0):
                        continue
                    sp_df["Date"] = pd.to_datetime(sp_df["Date"])

                    # Get the row with the latest date
                    latest_row = sp_df.loc[sp_df["Date"].idxmax()]

                    # Update the date in the latest row copy
                    new_row = latest_row.copy()
                    new_row["Date"] = pd.to_datetime(current_date)
                    new_row["Sales"] = np.nan
                    new_row["Purchase"] = np.nan
                    new_row["Net Fund"] = np.nan

                    # Append the updated row to the DataFrame
                    sales_purchase_dict[key] = sp_df._append(new_row, ignore_index=True)
                break

            if close_price and (not pd.isna(close_price)):
                closing_price = float(close_price)
                ed_closing_price = float(close_price)
                temp_df.at[index, 'CMP'] = ed_closing_price
                temp_df.at[index, 'CMP '] = closing_price
                temp_df.at[index, 'MV'] = closing_price * row['No. ']
                total_market_value += temp_df.at[index, 'MV']

                drill_down_df = enter_track(drill_down_df, current_date, symbol, row["Broker"], row["File"], row['No. '], row['Cost/Sh'], closing_price)

                if (sell_date != pd.to_datetime(current_date)):
                    values["Overall"] += closing_price * row['No. ']
                    values[row['Broker']] += closing_price * row['No. ']

                    if (dop == pd.to_datetime(current_date)):
                        purchase["Overall"] += row["Net Cost"]
                        purchase[row['Broker']] += row['Net Cost']

                else:
                    sales["Overall"] += row['Net Sale']
                    sales[row['Broker']] += row['Net Sale']

            else:
                print(f"No data found for {symbol_ns} on {current_date}")

    for keys in net_fund:
        if (not values[keys]) or  values[keys] == 0:
            continue
        net_fund[keys] = purchase[keys] - sales[keys]
        if not sales_purchase_dict[keys].empty:
            # Safely access the last row's 'Units' and 'NAV'
            prev_unit = sales_purchase_dict[keys].iloc[-1].Units
            prev_nav = sales_purchase_dict[keys].iloc[-1].NAV
        else:
            # Use fallback values when the DataFrame is empty
            prev_unit = values[keys] / 1000
            prev_nav = 1000  # Default value

        # Update 'units' calculation
        units[keys] = (prev_unit + (net_fund[keys] / prev_nav)) if not sales_purchase_dict[keys].empty else values[keys] / prev_nav

        # Update 'nav' calculation
        nav[keys] = values[keys] / units[keys]


    for key in sales_purchase_dict:
        if values[key] == 0:
            continue

        sales_purchase_dict[key].loc[len(sales_purchase_dict[key])] = [current_date ,values[key], purchase[key], sales[key], net_fund[key], units[key], nav[key]]



    current_date += timedelta(days=1)




# write the temp_df inside "Excel Automation.xlsx" file in Sheet1
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    temp_df.to_excel(writer, sheet_name=sheet_name, index=False)

save_sales_purchase_dict(sales_purchase_dict=sales_purchase_dict)
save_drill_down_df(drill_down_df=drill_down_df)


print("unlisted shares", type(unlisted_shares), unlisted_shares)

print("run time", (datetime.now() - start_timestamp))