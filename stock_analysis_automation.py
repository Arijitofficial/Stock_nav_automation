import numpy as np
import pandas as pd
from datetime import datetime, timedelta
from dateutil.parser import parse
from Utils.sales_purchase_util import init_dict, save_sales_purchase_dict
from Utils.drill_down_util import enter_track, init_drill_down_df, save_drill_down_df
from down_close_price_data import create_stock_price_df
from Utils.symbol_change_handler import map_symbols
from openpyxl import load_workbook

class StockAnalysis:
    def __init__(self, file_path, sheet_name):
        print(file_path)
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.df = self.load_data()
        self.brokers = self.df['Broker'].unique()
        self.sales_purchase_dict = init_dict(file_path="./Excels/sales_purchase_data.xlsx", broker_names=self.df['Broker'].unique())
        self.drill_down_df = init_drill_down_df()
        self.start_date = None
        self.end_date = None
        self.close_price_df = None

    def load_data(self):
        try:
            df = pd.read_excel(self.file_path, self.sheet_name, engine='openpyxl')
        except Exception:
            print("File not found. Please provide a valid file path.")
            self.file_path = input("Enter the file name: ")
            self.sheet_name = input("Enter the sheet name: ")
            df = pd.read_excel(self.file_path, self.sheet_name, engine='openpyxl')

        df = df[(df["Cat"].str.lower() == 'normal') | (df["Cat"].str.lower() == 'others')]
        df["NSE Name "] = df.groupby("Name of Shares")["NSE Name "].transform(lambda x: x.fillna(method='ffill').fillna(method='bfill'))
        df["Broker"].fillna(value="unknown", inplace=True)
        df["Broker"] = df["Broker"].str.replace(r"(?i)\bone\b", "IIFL", regex=True)
        return df

    def input_dates(self):
        while True:
            if "Overall" in self.sales_purchase_dict and len(self.sales_purchase_dict["Overall"]) > 0:
                print("Enter 1 to pick the start date from sales purchase records\nEnter 2 to pick the start date manually", end="\nput choice: ")
                choice = input()
            else:
                choice = "2"

            if choice == "1" and "Overall" in self.sales_purchase_dict and len(self.sales_purchase_dict["Overall"]) > 0:
                self.start_date = str(self.sales_purchase_dict["Overall"]['Date'].iloc[-1])
            else:
                self.start_date = input("Enter start date (yyyy-mm-dd): ")

            self.end_date = input("Enter end date (yyyy-mm-dd): ")

            try:
                self.start_date = parse(self.start_date).date()
                self.end_date = parse(self.end_date).date()
                break
            except Exception as e:
                print(e)
                print("Please enter valid date in yyyy-mm-dd format")

    def fetch_close_price_data(self):
        self.df['S. Date'] = pd.to_datetime(self.df['S. Date'], errors='coerce').dt.date
        filtered_df = self.df[(self.df['S. Date'].isnull()) | (self.df['S. Date'] > self.start_date)]

        unique_df = filtered_df[['NSE Name ', 'Symbol', 'Name of Shares']].drop_duplicates()
        unique_df['ticker'] = unique_df.apply(
            lambda row: f"{row['NSE Name ']}.NS" if row['Symbol'] == "NSE" else f"{row['NSE Name ']}.BO",
            axis=1
        )
        unique_df['bhav_keys'] = unique_df.apply(
            lambda row: f"{row['NSE Name ']}.NS" if row['Symbol'] == "NSE" else f"{row['Name of Shares']}.BO",
            axis=1
        )

        keys = list(set(unique_df['ticker']))
        symbols_dict = map_symbols(keys, start_date=str(self.start_date), end_date=str(self.end_date))
        self.close_price_df = create_stock_price_df(
            start_date=str(self.start_date), end_date=str(self.end_date), keys=keys, symbols_dict=symbols_dict
        )

    def process_data(self):
        current_date = self.start_date
        temp_df = self.df.copy()
        unlisted_shares = {}

        while current_date <= self.end_date:
            print(f"Processing date: {current_date}")
            total_market_value = 0
            brokers_n_overall = ["Overall"] + list(self.brokers)

            values = {broker: 0 for broker in brokers_n_overall}
            sales = {broker: 0 for broker in brokers_n_overall}
            purchase = {broker: 0 for broker in brokers_n_overall}
            net_fund = {broker: 0 for broker in brokers_n_overall}
            units = {broker: 0 for broker in brokers_n_overall}
            nav = {broker: 0 for broker in brokers_n_overall}


            for index, row in temp_df.iterrows():
                symbol = row['NSE Name ']
                dop = pd.to_datetime(row['DOP'])
                sell_date = pd.to_datetime(row['S. Date']) if pd.notnull(row['S. Date']) else None

                if pd.notnull(symbol) and (dop <= pd.to_datetime(current_date)) and (sell_date is None or pd.to_datetime(current_date) <= sell_date):
                    symbol_ns = symbol + (".NS" if row['Symbol'] == "NSE" else ".BO")

                    try:
                        if symbol_ns in self.close_price_df['ticker'].values:
                            date_column = current_date.strftime("%Y-%m-%d")

                            if date_column in self.close_price_df.columns:
                                close_price = self.close_price_df.loc[self.close_price_df['ticker'] == symbol_ns, date_column].values[0]

                                if pd.isna(close_price) or pd.isnull(close_price):
                                    close_price = row['Cost/Sh']
                            else:
                                raise ValueError(f"Date {date_column} not found in the data.")
                        else:
                            close_price = None
                    except Exception as e:
                        print(f"Error fetching data for {symbol_ns} on {current_date}: {e}")
                        date = pd.to_datetime(current_date)
                        if (date.dayofweek == 5) or (date.dayofweek == 6):
                            print("weekend day", current_date)

                        if len(self.sales_purchase_dict["Overall"]) <= 0:
                            break
                        print(current_date, e)
                        for key, sp_df in self.sales_purchase_dict.items():
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
                            self.sales_purchase_dict[key] = sp_df._append(new_row, ignore_index=True)
                        break

                    if close_price and (not pd.isna(close_price)):
                        closing_price = float(close_price)
                        temp_df.at[index, 'CMP'] = closing_price
                        temp_df.at[index, 'MV'] = closing_price * row['No. ']
                        total_market_value += temp_df.at[index, 'MV']

                        prev_value = values[row['Broker']]

                        if sell_date != pd.to_datetime(current_date):
                            values["Overall"] += closing_price * row['No. ']
                            values[row['Broker']] += closing_price * row['No. ']
                            if dop == pd.to_datetime(current_date):
                                purchase["Overall"] += row["Net Cost"]
                                purchase[row['Broker']] += row["Net Cost"]
                        else:
                            sales["Overall"] += row['Net Sale']
                            sales[row['Broker']] += row['Net Sale']
                        
                        current_mv = values[row['Broker']]
                        self.drill_down_df = enter_track(self.drill_down_df, current_date, symbol, row['Broker'], row['File'], row['No. '], row['Cost/Sh'], closing_price, current_mv)
                    else:
                        print(f"No data found for {symbol_ns} on {current_date}")

            for broker in values:
                isPurchaseOrSellHappened = not (sales[broker].empty and purchase[broker].empty)
                if(isPurchaseOrSellHappened and broker=='Sanctum' and values[broker] == 0):
                    print("sales puchase happened but value 0", current_date)
                if not ((not values[broker]) or  values[broker] == 0):
                    net_fund[broker] = purchase[broker] - sales[broker]
                    prev_unit = self.sales_purchase_dict[broker].iloc[-1].Units if not self.sales_purchase_dict[broker].empty else values[broker] / 1000
                    prev_nav = self.sales_purchase_dict[broker].iloc[-1].NAV if not self.sales_purchase_dict[broker].empty else 1000
                    units[broker] = (prev_unit + (net_fund[broker] / prev_nav)) if not self.sales_purchase_dict[broker].empty else values[broker] / prev_nav
                    nav[broker] = values[broker] / units[broker]

                if values[broker] == 0:
                    continue
                self.sales_purchase_dict[broker].loc[len(self.sales_purchase_dict[broker])] = [current_date, values[broker], purchase[broker], sales[broker], net_fund[broker], units[broker], nav[broker]]

            current_date += timedelta(days=1)

        with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            temp_df.to_excel(writer, sheet_name=self.sheet_name, index=False)

        save_sales_purchase_dict(self.sales_purchase_dict)
        save_drill_down_df(self.drill_down_df)

    def run(self):
        self.input_dates()
        self.fetch_close_price_data()
        self.process_data()

# Example usage:
if __name__ == "__main__":
    file_path = r"C:\Users\ASUS\Downloads\StockHolding_14.02.2025.xlsx"
    sheet_name = "Sheet1"
    stock_analysis = StockAnalysis(file_path, sheet_name)
    stock_analysis.run()
