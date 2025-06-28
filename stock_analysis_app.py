from datetime import timedelta
import json
import os
import sys
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from dateutil.parser import parse
from Utils.sales_purchase_util import init_dict, save_sales_purchase_dict
from Utils.drill_down_util import enter_track, init_drill_down_df, save_drill_down_df
from down_close_price_data import create_stock_price_df
from Utils.symbol_change_handler import map_symbols
import threading

class StockAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Stock Analysis Tool")
        self.root.geometry("600x500")

        # Variables
        self.config_file = "defaults.json"
        self.default_values = self.load_defaults()

        # Initialize class variables
        self.file_path = self.default_values.get("main_file_path", "")
        self.sales_purchase_file_path = self.default_values.get("sales_purchase_file_path", "")
        self.sheet_name = self.default_values.get("sheet_name", "")
        self.start_date = None
        self.end_date = None
        self.df = None
        self.sales_purchase_dict = None
        self.drill_down_df = None
        self.close_price_df = None
        self.brokers = []

        # Start the app with the file input screen
        self.file_input_and_sheet_name_screen()

    def clear_screen(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def load_defaults(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, "r") as file:
                return json.load(file)
        return {}

    def save_defaults(self):
        defaults = {
            "main_file_path": self.file_path,
            "sales_purchase_file_path": self.sales_purchase_file_path,
            "sheet_name": self.sheet_name_var.get(),
        }
        with open(self.config_file, "w") as file:
            json.dump(defaults, file)


    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_path = file_path
            self.file_path_label.config(text=file_path)
            self.save_defaults()

    def browse_sales_purchase_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.sales_purchase_file_path = file_path
            self.sales_purchase_file_path_label.config(text=file_path)
            self.save_defaults()

    def clear_inputs(self):
        self.file_path = ""
        self.sales_purchase_file_path = ""
        self.sheet_name_var.set("")
        self.file_path_label.config(text="No file selected")
        self.sales_purchase_file_path_label.config(text="No file selected")
        self.save_defaults()  # Save empty values

    def file_input_and_sheet_name_screen(self):
        self.clear_screen()

        # Input for main data file
        ttk.Label(self.root, text="Select Main Excel File:").grid(row=0, column=0, pady=10, padx=10, sticky='w')
        self.file_path_label = ttk.Label(self.root, text=self.file_path or "No file selected", width=40, anchor="w")
        self.file_path_label.grid(row=0, column=1, pady=10, padx=10)
        ttk.Button(self.root, text="Browse", command=self.browse_file).grid(row=0, column=2, pady=10, padx=10)

        # Input for sales/purchase file (optional)
        ttk.Label(self.root, text="Select Sales/Purchase File (Optional):").grid(row=1, column=0, pady=10, padx=10, sticky='w')
        self.sales_purchase_file_path_label = ttk.Label(self.root, text=self.sales_purchase_file_path or "No file selected", width=40, anchor="w")
        self.sales_purchase_file_path_label.grid(row=1, column=1, pady=10, padx=10)
        ttk.Button(self.root, text="Browse", command=self.browse_sales_purchase_file).grid(row=1, column=2, pady=10, padx=10)

        # Input for sheet name
        ttk.Label(self.root, text="Enter Sheet Name:").grid(row=2, column=0, pady=10, padx=10, sticky='w')
        self.sheet_name_var = tk.StringVar(value=self.sheet_name)
        ttk.Entry(self.root, textvariable=self.sheet_name_var, width=30).grid(row=2, column=1, pady=10, padx=10)

        # Buttons (Aligned to the right)
        # Configure grid to ensure alignment
        self.root.grid_columnconfigure(0, weight=1)  # Allows left side to expand
        self.root.grid_columnconfigure(1, weight=1)  # Center area
        self.root.grid_columnconfigure(2, weight=1)  # Rightmost area

        # Buttons (Aligned to the right)
        button_frame = ttk.Frame(self.root)
        button_frame.grid(row=3, column=0, columnspan=3, pady=20, padx=10, sticky="e")

        
        ttk.Button(button_frame, text="Next", command=self.date_input_screen).pack(side="right", padx=5)
        ttk.Button(button_frame, text="Clear Selection", command=self.clear_inputs).pack(side="right")

    def date_input_screen(self):
        self.save_defaults()
        self.sheet_name = self.sheet_name_var.get()
        if not self.file_path:
            messagebox.showerror("Error", "Please select the main Excel file.")
            return

        if not self.sheet_name:
            messagebox.showerror("Error", "Please enter a sheet name.")
            return

        # Load data and initialize dicts
        valid_files = self.load_data()
        if not valid_files:
            return

        self.clear_screen()

        # Input for dates
        ttk.Label(self.root, text="Start Date (optional, yyyy-mm-dd):").pack(pady=10)
        self.start_date_var = tk.StringVar()

        # Pre-fill start date if sales purchase dict is not empty
        if self.sales_purchase_dict and "Overall" in self.sales_purchase_dict and not self.sales_purchase_dict["Overall"].empty:
            self.start_date_var.set(str(self.sales_purchase_dict["Overall"]["Date"].iloc[-1]))

        ttk.Entry(self.root, textvariable=self.start_date_var, width=20).pack(pady=5)

        ttk.Label(self.root, text="End Date (yyyy-mm-dd):").pack(pady=10)
        self.end_date_var = tk.StringVar()
        ttk.Entry(self.root, textvariable=self.end_date_var, width=20).pack(pady=5)

        # buttons
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=20)
        ttk.Button(button_frame, text="Previous", command=self.file_input_and_sheet_name_screen).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Next", command=self.fetch_close_price_data_screen).pack(side=tk.LEFT)


    def fetch_close_price_data_screen(self):
        self.start_date = self.start_date_var.get()
        self.end_date = self.end_date_var.get()

        if not self.end_date:
            messagebox.showerror("Error", "End date is required.")
            return

        # Validate and parse dates
        try:
            self.start_date = parse(self.start_date).date() if self.start_date else None
            self.end_date = parse(self.end_date).date()
        except Exception as e:
            messagebox.showerror("Error", f"Invalid date format: {e}")
            return

        self.run()

    def browse_file(self):
        self.file_path = filedialog.askopenfilename(
            title="Select Main Excel File", filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if self.file_path:
            self.file_path_label.config(text=self.file_path)


    def load_data(self):
        try:
            self.df = pd.read_excel(self.file_path, self.sheet_name, engine='openpyxl')
            self.df = self.df[(self.df["Cat"].str.lower() == 'normal') | (self.df["Cat"].str.lower() == 'others')]
            self.df["NSE Name "] = self.df.groupby("Name of Shares")["NSE Name "].transform(
                lambda x: x.fillna(method='ffill').fillna(method='bfill')
            )
            self.df["Broker"].fillna(value="unknown", inplace=True)
            self.df["Broker"] = self.df["Broker"].str.replace(r"(?i)\bone\b", "IIFL", regex=True)
            self.brokers = self.df["Broker"].unique()
            print(self.brokers)
            self.sales_purchase_dict = init_dict(
                file_path=self.sales_purchase_file_path, 
                broker_names=self.brokers
            )
            self.drill_down_df = init_drill_down_df()
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {e}")
            return False

    def fetch_close_price_data(self):
        try:
            # Show popup
            self.show_loading_popup()

            # Fetching process
            self.df['S. Date'] = pd.to_datetime(self.df['S. Date'], errors='coerce').dt.date
            filtered_df = self.df[(self.df['S. Date'].isnull()) | (self.df['S. Date'] > self.start_date)] if self.start_date else self.df

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
                start_date=str(self.start_date), 
                end_date=str(self.end_date), 
                keys=keys, 
                symbols_dict=symbols_dict
            )
            self.close_price_df.to_csv("close_prices_dataframe.csv", index=False)
            # Close popup after completion
            self.loading_popup.destroy()
            self.process_data()

        except Exception as e:
            self.loading_popup.destroy()
            messagebox.showerror("Error", f"Failed to fetch close price data: {e}")

    # Function to create the loading popup
    def show_loading_popup(self):
        self.loading_popup = tk.Toplevel(self.root)
        self.loading_popup.title("Loading")
        self.loading_popup.geometry("250x100")
        self.loading_popup.resizable(False, False)
        tk.Label(self.loading_popup, text="Downloading...", font=("Arial", 12)).pack(pady=20)
        self.loading_popup.grab_set()  # Keeps it on top
        self.root.update_idletasks()

    # Function to run `fetch_close_price_data` in a separate thread
    def fetch_data_threaded(self):
        threading.Thread(target=self.fetch_close_price_data, daemon=True).start()
    
    def process_data(self):
        # Progress window setup
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Processing Data")
        progress_window.geometry("300x100")
        ttk.Label(progress_window, text="Processing data...").pack(pady=10)
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=200, mode="determinate")
        progress_bar.pack(pady=10)
        

        total_days = (self.end_date - self.start_date).days + 1
        progress_bar["maximum"] = total_days

        def process():
            update_frequency = 5  # Update progress bar every 10 days
            days_processed = 0
            current_date = self.start_date
            temp_df = self.df.copy()
            unlisted_shares = {}
            
            while current_date <= self.end_date:
                days_processed += 1
                is_sp_handled = False
                if days_processed % update_frequency == 0 or current_date == self.end_date:
                    progress_bar["value"] = days_processed
                    progress_window.update_idletasks()
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
                            is_sp_handled = True
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
                            self.drill_down_df = enter_track(self.drill_down_df, current_date, symbol, row['Broker'], row['File'], row['No. '], row['Cost/Sh'], closing_price, current_mv - prev_value)
                        else:
                            print(f"No data found for {symbol_ns} on {current_date}")

                for broker in values:
                    # todo: sell handle even for value 0
                    isPurchaseOrSellHappened = (sales[broker]!=0 or purchase[broker]!=0)
                    if isPurchaseOrSellHappened or (not ((not values[broker]) or  values[broker] == 0)):
                        net_fund[broker] = purchase[broker] - sales[broker]
                        prev_unit = self.sales_purchase_dict[broker].iloc[-1].Units if not self.sales_purchase_dict[broker].empty else values[broker] / 1000
                        prev_nav = self.sales_purchase_dict[broker].iloc[-1].NAV if not self.sales_purchase_dict[broker].empty else 1000
                        prev_value = self.sales_purchase_dict[broker].iloc[-1].Value if not self.sales_purchase_dict[broker].empty else 0
                        
                        if(prev_nav == 0):
                            print("prev_value: ", prev_value)
                            prev_nav = 1000
                        
                        units[broker] = (prev_unit + (net_fund[broker] / prev_nav)) if (int(prev_value)!=0) else values[broker] / prev_nav
                        nav[broker] = values[broker] / units[broker] if units[broker]!=0 else 0
                        if(values[broker]==0):
                            units[broker] = 0
                            nav[broker] = 0
                    # todo: sell handle even for value 0
                    date = pd.to_datetime(current_date)
                    if (not isPurchaseOrSellHappened) and values[broker] == 0 and (is_sp_handled):
                        continue

                    self.sales_purchase_dict[broker].loc[len(self.sales_purchase_dict[broker])] = [current_date, values[broker], purchase[broker], sales[broker], net_fund[broker], units[broker], nav[broker]]

                current_date += timedelta(days=1)
            
            progress_window.destroy()

            with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                temp_df.to_excel(writer, sheet_name=self.sheet_name, index=False)
            
            if self.root is not None and self.root:
                self.on_close()

            save_sales_purchase_dict(self.sales_purchase_dict)
            save_drill_down_df(self.drill_down_df)

        threading.Thread(target=process).start()

    def on_close(self):
        self.clear_screen()
        messagebox.showinfo("Success", "Process completed successfully!")
        self.root.after(5000, self.terminate_process)
        
    def terminate_process(self):
        # Terminate the entire Python process
        sys.exit(0)

    def run(self):
        self.fetch_data_threaded()
        
        # Further processing would go here


# Run the app
if __name__ == "__main__":
    root = tk.Tk()
    app = StockAnalysisApp(root)
    root.mainloop()


