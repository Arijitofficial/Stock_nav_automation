import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import yfinance as yf
from dateutil.parser import parse
from Utils.sales_purchase_util import init_dict, save_sales_purchase_dict
from Utils.drill_down_util import enter_track, init_drill_down_df, save_drill_down_df
import openpyxl
from openpyxl import load_workbook



# Define the file containing the records

file_path = './Excels/Excel Automation.xlsx'
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



cp_file_path = "./Excels/nse_close_price_data.csv"
close_price_df = pd.read_csv(file_path)


# iterate over start to end dates and fetch current market prices and update files



# Initialize current_date
current_date = start_date
temp_df = df.copy()
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
            symbol_ns = symbol + ".NS"  # Append ".NS" to the symbol name for Yahoo Finance
            try:
                print("stock found in date range", symbol_ns)
                stock_data = yf.download(symbol_ns, start=current_date, end=current_date + timedelta(days=1))

            except Exception as e:
                print(f"Error fetching data for {symbol_ns} on {current_date}: {e}")
                date = pd.to_datetime(current_date)
                if (date.dayofweek == 5) or (date.dayofweek == 6):
                    print("weekend day", current_date)

                # Iterate through each DataFrame in the dictionary
                for key, df in sales_purchase_dict.items():
                    # Ensure date column is in datetime format
                    df["date"] = pd.to_datetime(df["date"])

                    # Get the row with the latest date
                    latest_row = df.loc[df["date"].idxmax()]

                    # Update the date in the latest row copy
                    new_row = latest_row.copy()
                    new_row["date"] = pd.to_datetime(current_date)
                    new_row["sales"] = np.nan
                    new_row["purchase"] = np.nan
                    new_row["net_fund"] = np.nan

                    # Append the updated row to the DataFrame
                    sales_purchase_dict[key] = df.append(new_row, ignore_index=True)
                break

            if not stock_data.empty:
                print("stock found in yahoo", stock_data.iloc[0]['Close'] * 1000)
                closing_price = float(stock_data.iloc[0]['Close'])
                ed_closing_price = float(stock_data.iloc[0]['Close'])
                # cmp added here
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
        if values[keys] == 0:
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
        print("keys in dictionary", key)
        print("values",  [current_date ,values[key], purchase[key], sales[key], net_fund[key], units[key], nav[key]])
        if values[key] == 0:
            continue

        sales_purchase_dict[key].loc[len(sales_purchase_dict[key])] = [current_date ,values[key], purchase[key], sales[key], net_fund[key], units[key], nav[key]]



    current_date += timedelta(days=1)




# write the temp_df inside "Excel Automation.xlsx" file in Sheet1
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    temp_df.to_excel(writer, sheet_name=sheet_name, index=False)

save_sales_purchase_dict(sales_purchase_dict=sales_purchase_dict)
save_drill_down_df(drill_down_df=drill_down_df)